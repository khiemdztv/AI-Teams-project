/**
 * Fork addition: a smaller JPEG copy of a large image, so Claude can look at
 * it (its image input stops at 5MB; a phone photo or a screenshot of a big
 * screen is often more).
 *
 * Windows: a fixed PowerShell script using System.Drawing, which ships with
 * Windows. Paths reach it through environment variables, never through the
 * script text, so a file name cannot change what runs.
 *
 * Elsewhere: ImageMagick 7 (`magick`), given argv directly with no shell.
 * Paths are absolute, so none can be read as an option. Only the first frame
 * of `src` is read (`[0]`), so an animated GIF or a multi-page TIFF stays one
 * picture.
 */

import { existsSync, statSync } from 'fs'
import { isAbsolute } from 'path'

const SCRIPT = `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$img = [System.Drawing.Image]::FromFile($env:PREVIEW_SRC)
try {
  $scale = [Math]::Min(1.0, [double]$env:PREVIEW_MAX_SIDE / [Math]::Max($img.Width, $img.Height))
  $w = [Math]::Max(1, [int]($img.Width * $scale)); $h = [Math]::Max(1, [int]($img.Height * $scale))
  $bmp = New-Object System.Drawing.Bitmap $w, $h
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.Clear([System.Drawing.Color]::White)
  $g.InterpolationMode = 'HighQualityBicubic'
  $g.DrawImage($img, 0, 0, $w, $h)
  $enc = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
  $p = New-Object System.Drawing.Imaging.EncoderParameters 1
  $p.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality), ([long]80)
  $bmp.Save($env:PREVIEW_DST, $enc, $p)
  $g.Dispose(); $bmp.Dispose()
} finally { $img.Dispose() }
`

/** Write a JPEG copy of `src`, longest side at most `maxSide`, to `dst`. Returns `dst`. */
export async function makeImagePreview(src: string, dst: string, maxSide = 2000): Promise<string> {
  const proc =
    process.platform === 'win32'
      ? Bun.spawn(['powershell.exe', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', SCRIPT], {
          env: { ...process.env, PREVIEW_SRC: src, PREVIEW_DST: dst, PREVIEW_MAX_SIDE: String(maxSide) },
          stdin: 'ignore',
          stdout: 'ignore',
          stderr: 'pipe',
        })
      : spawnMagick(src, dst, maxSide)
  const timer = setTimeout(() => proc.kill(), 60_000)
  const code = await proc.exited
  clearTimeout(timer)
  if (code !== 0 || !existsSync(dst) || statSync(dst).size === 0) {
    const err = (await new Response(proc.stderr).text()).split('\n').find(l => l.trim()) ?? `exit ${code}`
    throw new Error(`could not make a smaller copy: ${err.trim().slice(0, 200)}`)
  }
  return dst
}

function spawnMagick(src: string, dst: string, maxSide: number) {
  if (!isAbsolute(src) || !isAbsolute(dst)) throw new Error('image preview paths must be absolute')
  const side = Math.max(1, Math.floor(maxSide))
  return Bun.spawn(
    [
      'magick',
      `${src}[0]`,
      '-auto-orient',
      '-resize',
      `${side}x${side}>`,
      '-background',
      'white',
      '-flatten',
      '-quality',
      '80',
      `jpg:${dst}`,
    ],
    { stdin: 'ignore', stdout: 'ignore', stderr: 'pipe' },
  )
}
