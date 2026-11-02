/**
 * make_pdf (fork addition): what Claude passes never becomes markup, images
 * come only from the allowed folders, and the page allows no network or file
 * access. The last test prints a real PDF when Edge is installed.
 */

import { describe, expect, test, beforeEach } from 'bun:test'
import { mkdtempSync, mkdirSync, writeFileSync, readFileSync, existsSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { buildPdfHtml, pdfFileName, printHtmlToPdf, findEdge } from '../src/pdf.js'

const PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
)

let dir: string
let inbox: string
let exportDir: string
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'pdf-'))
  inbox = join(dir, 'inbox')
  exportDir = join(dir, 'bot-export')
  mkdirSync(inbox)
  mkdirSync(exportDir)
})

describe('page', () => {
  test('text is escaped, so it cannot add tags, iframes or links', () => {
    const html = buildPdfHtml(
      { title: '<script>x</script>', text: '<iframe src="file:///C:/x/.env"></iframe>\n<img src=https://evil>' },
      [inbox],
    )
    expect(html).not.toContain('<iframe')
    expect(html).not.toContain('<script>x')
    expect(html).not.toContain('<img src=https')
    expect(html).toContain('&lt;iframe')
  })

  test('the page allows no network, script or file access', () => {
    const html = buildPdfHtml({ text: 'a' }, [inbox])
    expect(html).toContain(`content="default-src 'none'; img-src data:; style-src 'unsafe-inline'"`)
  })

  test('headings, bullets, paragraphs and bold', () => {
    const html = buildPdfHtml({ text: '# Tiêu đề\nDòng 1\nDòng 2\n\n- một\n- **hai**' }, [inbox])
    expect(html).toContain('<h2>Tiêu đề</h2>')
    expect(html).toContain('<p>Dòng 1<br>Dòng 2</p>')
    expect(html).toContain('<ul><li>một</li><li><strong>hai</strong></li></ul>')
  })

  test('an image from the inbox is embedded as a data URI', () => {
    const photo = join(inbox, 'a.png')
    writeFileSync(photo, PNG)
    expect(buildPdfHtml({ images: [photo] }, [inbox, exportDir])).toContain('src="data:image/png;base64,')
  })

  test('an image outside the allowed folders is refused', () => {
    const outside = join(dir, 'secret.png')
    writeFileSync(outside, PNG)
    expect(() => buildPdfHtml({ images: [outside] }, [inbox, exportDir])).toThrow(/not in an allowed folder/)
  })

  test('a file that is not an image is refused', () => {
    const text = join(inbox, 'x.png')
    writeFileSync(text, 'MSTEAMS_APP_PASSWORD=...')
    expect(() => buildPdfHtml({ images: [text] }, [inbox])).toThrow(/not a PNG/)
  })

  test('an empty PDF is refused', () => {
    expect(() => buildPdfHtml({}, [inbox])).toThrow(/nothing/)
  })
})

describe('file name', () => {
  test('a bare name in the export folder, ending in .pdf', () => {
    expect(pdfFileName('bao-cao')).toBe('bao-cao.pdf')
    expect(pdfFileName('Test 2.PDF')).toBe('Test 2.PDF')
    expect(pdfFileName('..\\..\\Windows\\x.pdf')).toBe('x.pdf')
    expect(pdfFileName('a/b:c.pdf')).toBe('b_c.pdf')
    expect(() => pdfFileName('')).toThrow()
  })
})

describe.if(!!findEdge())('Edge / Chrome', () => {
  test('prints a real PDF', async () => {
    const photo = join(inbox, 'a.png')
    writeFileSync(photo, PNG)
    const out = join(exportDir, 'test.pdf')
    await printHtmlToPdf(buildPdfHtml({ title: 'Thử', text: 'Tiếng Việt có dấu', images: [photo] }, [inbox]), out)
    expect(existsSync(out)).toBe(true)
    expect(readFileSync(out).subarray(0, 5).toString()).toBe('%PDF-')
  }, 60_000)
})
