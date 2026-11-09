import { describe, expect, test } from 'bun:test'
import { mkdtempSync, writeFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import {
  MAX_SLIDE_SECONDS,
  MAX_VIDEO_SLIDES,
  MIN_SLIDE_SECONDS,
  buildSlideHtml,
  ffmpegArgs,
  makeSlideVideo,
  normalizeSlides,
  videoFileName,
} from '../src/video.js'

const PNG = Buffer.from('89504e470d0a1a0a0000000d4948445200000001000000010806000000', 'hex')

describe('slide video (fork, 09/11)', () => {
  test('the file name is a bare .mp4 name', () => {
    expect(videoFileName('gioi-thieu')).toBe('gioi-thieu.mp4')
    expect(videoFileName('../../x.MP4')).toBe('x.MP4')
    expect(videoFileName('..\\..\\x')).toBe('x.mp4')
    expect(() => videoFileName('')).toThrow(/required/)
  })

  test('each slide gets a duration from how much there is to read', () => {
    const [short, long, set] = normalizeSlides([
      { title: 'Xin chào' },
      { text: 'từ '.repeat(60) },
      { title: 'x', seconds: 999 },
    ])
    expect(short.seconds).toBeGreaterThanOrEqual(MIN_SLIDE_SECONDS)
    expect(long.seconds).toBeGreaterThan(short.seconds)
    expect(set.seconds).toBe(MAX_SLIDE_SECONDS)
  })

  test('empty, too many or too wordy slides are refused', () => {
    expect(() => normalizeSlides([])).toThrow(/non-empty/)
    expect(() => normalizeSlides('x')).toThrow(/non-empty/)
    expect(() => normalizeSlides([{ title: '  ' }])).toThrow(/empty/)
    expect(() => normalizeSlides(Array(MAX_VIDEO_SLIDES + 1).fill({ title: 'x' }))).toThrow(/at most/)
    expect(() => normalizeSlides([{ text: 'x'.repeat(5000) }])).toThrow(/split it/)
  })

  test('slide text is escaped and the page can fetch nothing', () => {
    const html = buildSlideHtml({ title: '<script>x</script>', caption: '<img src=http://e.com>' }, [])
    expect(html).not.toContain('<script>x')
    expect(html).not.toContain('<img src=http')
    expect(html).toContain("default-src 'none'")
  })

  test('images come only from the allowed folders', () => {
    const root = mkdtempSync(join(tmpdir(), 'video-'))
    const inside = join(root, 'a.png')
    writeFileSync(inside, PNG)
    expect(buildSlideHtml({ image: inside }, [root])).toContain('data:image/png;base64,')
    const outside = join(mkdtempSync(join(tmpdir(), 'video-out-')), 'b.png')
    writeFileSync(outside, PNG)
    expect(() => buildSlideHtml({ image: outside }, [root])).toThrow()
  })

  test('each fade starts where the slides before it end', () => {
    const args = ffmpegArgs(['a.png', 'b.png', 'c.png'], [4, 5, 3], 'out.mp4')
    const graph = args[args.indexOf('-filter_complex') + 1]
    expect(graph).toContain('xfade=transition=fade:duration=0.5:offset=3.50[x1]')
    expect(graph).toContain('[x1][v2]xfade=transition=fade:duration=0.5:offset=8.00[x2]')
    expect(args[args.indexOf('-map') + 1]).toBe('[x2]')
    expect(args.at(-1)).toBe('out.mp4')
  })

  test('one slide needs no fade', () => {
    const args = ffmpegArgs(['a.png'], [4], 'out.mp4')
    expect(args[args.indexOf('-filter_complex') + 1]).not.toContain('xfade')
    expect(args[args.indexOf('-map') + 1]).toBe('[v0]')
  })

  test('a missing ffmpeg is said plainly, after the slides were checked', async () => {
    const out = join(mkdtempSync(join(tmpdir(), 'video-')), 'x.mp4')
    await expect(makeSlideVideo([{ title: 'x' }], out, [], { chrome: '/bin/true', ffmpeg: undefined })).rejects.toThrow(
      /ffmpeg/,
    )
    await expect(makeSlideVideo([], out, [], { chrome: undefined, ffmpeg: undefined })).rejects.toThrow(/non-empty/)
  })
})
