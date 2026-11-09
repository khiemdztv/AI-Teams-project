/**
 * Fork addition (09/11): a slide video (MP4) from text and images, without
 * giving Claude a shell.
 *
 * Same shape as src/pdf.ts, whose page rules this reuses: Claude passes
 * structured slides, never markup. Each slide becomes one HTML page (text
 * escaped, images read here from the allowed folders and embedded as data
 * URIs, a CSP that allows nothing but inline styles and data: images, every
 * host name failing to resolve), headless Chrome screenshots it, and ffmpeg
 * joins the pictures with a short fade between slides. No sound.
 */

import { existsSync, mkdtempSync, rmSync, statSync, copyFileSync, writeFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { pathToFileURL } from 'url'
import { embedImage, escapeHtml, findEdge, textToHtml } from './pdf.js'

export const MAX_VIDEO_SLIDES = 30
export const MAX_SLIDE_SECONDS = 30
export const MIN_SLIDE_SECONDS = 2
const MAX_SLIDE_TEXT_CHARS = 1200
const WIDTH = 1280
const HEIGHT = 720
const FADE_SECONDS = 0.5
const FPS = 30
const SHOT_TIMEOUT_MS = 30_000
const ENCODE_TIMEOUT_MS = 5 * 60_000

export type Slide = {
  title?: string
  /** As make_pdf: blank lines separate paragraphs, "- " starts a bullet, **bold**. */
  text?: string
  /** One absolute image path inside an allowed folder. */
  image?: string
  /** A subtitle line in a bar along the bottom. */
  caption?: string
  /** How long the slide stays up; default from how much there is to read. */
  seconds?: number
}

/** A file name for the export folder: a bare name ending in .mp4. */
export function videoFileName(raw: string): string {
  let name = (String(raw ?? '').trim().split(/[\\/]/).pop() ?? '').replace(/[<>:"/\\|?*\x00-\x1f]/g, '_')
  if (!name || name === '.mp4' || /^\.+$/.test(name)) throw new Error('file_name is required')
  if (!/\.mp4$/i.test(name)) name += '.mp4'
  if (name.length > 120) throw new Error('file_name is too long')
  return name
}

/** Check the slides Claude passed and fill in each one's duration. */
export function normalizeSlides(raw: unknown): (Slide & { seconds: number })[] {
  if (!Array.isArray(raw) || raw.length === 0) throw new Error('slides must be a non-empty array')
  if (raw.length > MAX_VIDEO_SLIDES) throw new Error(`at most ${MAX_VIDEO_SLIDES} slides`)
  return raw.map((item, i) => {
    const s = (item ?? {}) as Record<string, unknown>
    const str = (k: string) => (typeof s[k] === 'string' && (s[k] as string).trim() ? (s[k] as string) : undefined)
    const slide: Slide = { title: str('title'), text: str('text'), image: str('image'), caption: str('caption') }
    if (!slide.title && !slide.text && !slide.image && !slide.caption) throw new Error(`slides[${i}] is empty`)
    if ((slide.text?.length ?? 0) > MAX_SLIDE_TEXT_CHARS) {
      throw new Error(`slides[${i}].text is longer than ${MAX_SLIDE_TEXT_CHARS} characters: split it`)
    }
    // About three words a second, plus time to take the slide in.
    const words = [slide.title, slide.text, slide.caption].join(' ').split(/\s+/).filter(Boolean).length
    const given = Number(s.seconds)
    const seconds = Number.isFinite(given) && given > 0 ? given : 2 + words / 3
    return { ...slide, seconds: Math.round(Math.min(MAX_SLIDE_SECONDS, Math.max(MIN_SLIDE_SECONDS, seconds)) * 10) / 10 }
  })
}

/** One slide as a 1280×720 page. */
export function buildSlideHtml(slide: Slide, imageRoots: string[]): string {
  const image = slide.image ? embedImage(slide.image, imageRoots) : ''
  const body = slide.text ? textToHtml(slide.text) : ''
  const layout = image && body ? 'split' : image ? 'picture' : 'words'
  return `<!doctype html>
<html><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'">
<style>
html, body { margin: 0; width: ${WIDTH}px; height: ${HEIGHT}px; overflow: hidden; background: #fff }
body { font-family: Carlito, Calibri, 'Noto Sans', 'DejaVu Sans', sans-serif; color: #0A2540; position: relative }
.frame { position: absolute; inset: 56px 72px ${slide.caption ? 120 : 56}px 72px; display: flex; flex-direction: column }
h1 { font-size: 46px; line-height: 1.15; margin: 0 0 12px }
.rule { width: 96px; height: 4px; background: #B1722A; margin: 0 0 28px; flex: none }
.main { flex: 1; min-height: 0; display: flex; gap: 40px; align-items: center }
.main.words { align-items: flex-start }
.words { flex: 1; font-size: 30px; line-height: 1.4 }
.words h2 { font-size: 34px; margin: 0 0 10px } .words h3 { font-size: 30px; margin: 0 0 8px }
.words p { margin: 0 0 14px } .words ul { margin: 0; padding-left: 1.1em } .words li { margin: 0 0 10px }
.pic { flex: 1; height: 100%; display: flex; align-items: center; justify-content: center }
.pic img { max-width: 100%; max-height: 100%; object-fit: contain }
.picture .pic { flex: 1 } .split .pic { flex: 0 0 46% }
.only-title h1 { font-size: 64px } .only-title { justify-content: center }
.caption { position: absolute; left: 0; right: 0; bottom: 0; min-height: 88px; box-sizing: border-box;
  padding: 18px 72px; background: #0A2540; color: #fff; font-size: 28px; line-height: 1.3;
  display: flex; align-items: center }
</style></head><body>
<div class="frame${!body && !image ? ' only-title' : ''}">
${slide.title ? `<h1>${escapeHtml(slide.title)}</h1>\n<div class="rule"></div>` : ''}
${body || image ? `<div class="main ${layout}">${body ? `<div class="words">${body}</div>` : ''}${image ? `<div class="pic"><img src="${image}"></div>` : ''}</div>` : ''}
</div>
${slide.caption ? `<div class="caption">${escapeHtml(slide.caption)}</div>` : ''}
</body></html>
`
}

/** ffmpeg arguments joining the slide pictures, a fade between each pair. */
export function ffmpegArgs(pictures: string[], seconds: number[], out: string): string[] {
  const args = ['-hide_banner', '-loglevel', 'error', '-y']
  pictures.forEach((p, i) => args.push('-loop', '1', '-framerate', String(FPS), '-t', String(seconds[i]), '-i', p))
  const filters: string[] = pictures.map((_, i) => `[${i}:v]scale=${WIDTH}:${HEIGHT},setsar=1,format=yuv420p[v${i}]`)
  let last = 'v0'
  let offset = 0
  for (let i = 1; i < pictures.length; i++) {
    offset += seconds[i - 1] - FADE_SECONDS
    const label = `x${i}`
    filters.push(`[${last}][v${i}]xfade=transition=fade:duration=${FADE_SECONDS}:offset=${offset.toFixed(2)}[${label}]`)
    last = label
  }
  args.push('-filter_complex', filters.join(';'), '-map', `[${last}]`)
  args.push('-c:v', 'libx264', '-preset', 'veryfast', '-crf', '23', '-pix_fmt', 'yuv420p', '-r', String(FPS))
  args.push('-movflags', '+faststart', out)
  return args
}

export function findFfmpeg(env: NodeJS.ProcessEnv = process.env): string | undefined {
  return [env.MSTEAMS_FFMPEG_PATH, '/usr/bin/ffmpeg', '/usr/local/bin/ffmpeg'].find(
    (p): p is string => !!p && existsSync(p),
  )
}

async function run(cmd: string[], timeoutMs: number): Promise<{ code: number; stderr: string }> {
  const proc = Bun.spawn(cmd, { stdin: 'ignore', stdout: 'ignore', stderr: 'pipe' })
  const timer = setTimeout(() => proc.kill(), timeoutMs)
  const [code, stderr] = await Promise.all([proc.exited, new Response(proc.stderr).text()])
  clearTimeout(timer)
  return { code, stderr }
}

/** Render `raw` slides to `outPath`. Returns the video's length in seconds. */
export async function makeSlideVideo(
  raw: unknown,
  outPath: string,
  imageRoots: string[],
  tools = { chrome: findEdge(), ffmpeg: findFfmpeg() },
): Promise<number> {
  const slides = normalizeSlides(raw)
  // Built before any process starts, so a bad image fails fast.
  const pages = slides.map(slide => buildSlideHtml(slide, imageRoots))
  if (!tools.chrome) throw new Error('making a video needs Google Chrome, which was not found')
  if (!tools.ffmpeg) throw new Error('making a video needs ffmpeg, which is not installed on the bot machine')

  const work = mkdtempSync(join(tmpdir(), 'msteams-video-'))
  try {
    const pictures: string[] = []
    for (const [i, html] of pages.entries()) {
      const page = join(work, `slide-${i}.html`)
      const picture = join(work, `slide-${i}.png`)
      writeFileSync(page, html, { mode: 0o600 })
      await run(
        [
          tools.chrome,
          '--headless=new',
          '--disable-gpu',
          '--no-first-run',
          '--no-default-browser-check',
          '--disable-extensions',
          '--disable-sync',
          '--hide-scrollbars',
          `--user-data-dir=${join(work, 'profile')}`,
          '--host-resolver-rules=MAP * ~NOTFOUND',
          `--window-size=${WIDTH},${HEIGHT}`,
          `--screenshot=${picture}`,
          pathToFileURL(page).href,
        ],
        SHOT_TIMEOUT_MS,
      )
      if (!existsSync(picture) || statSync(picture).size === 0) throw new Error(`slide ${i + 1} could not be drawn`)
      pictures.push(picture)
    }
    const video = join(work, 'out.mp4')
    const seconds = slides.map(s => s.seconds)
    const encoded = await run([tools.ffmpeg, ...ffmpegArgs(pictures, seconds, video)], ENCODE_TIMEOUT_MS)
    if (encoded.code !== 0 || !existsSync(video) || statSync(video).size === 0) {
      throw new Error(`ffmpeg failed: ${encoded.stderr.trim().split('\n').pop() ?? 'no output'}`)
    }
    copyFileSync(video, outPath)
    return Math.round((seconds.reduce((a, b) => a + b, 0) - FADE_SECONDS * (seconds.length - 1)) * 10) / 10
  } finally {
    try {
      rmSync(work, { recursive: true, force: true })
    } catch {
      // Chrome may still hold its profile for a moment; the temp folder is harmless.
    }
  }
}
