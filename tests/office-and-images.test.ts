/**
 * Fork additions: images pasted into a chat (fetched from the Bot Connector
 * with the bot token), and text extracted from Word / Excel / PowerPoint
 * attachments so Claude can Read them.
 */

import { describe, expect, test, beforeEach, afterEach } from 'bun:test'
import { mkdtempSync, readFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { deflateRawSync } from 'zlib'
import { AttachmentHandles, sniffImageExt } from '../src/attachments.js'
import { readZip, docxToText, xlsxToText, pptxToText, officeToText, isOfficeFile, decodeXml } from '../src/office-text.js'

/** A minimal zip writer for fixtures. CRCs are left 0: readZip does not check them. */
function zip(files: Record<string, string>): Buffer {
  const locals: Buffer[] = []
  const centrals: Buffer[] = []
  let offset = 0
  for (const [name, content] of Object.entries(files)) {
    const nameBuf = Buffer.from(name, 'utf8')
    const raw = Buffer.from(content, 'utf8')
    const data = deflateRawSync(raw)
    const local = Buffer.alloc(30)
    local.writeUInt32LE(0x04034b50, 0)
    local.writeUInt16LE(8, 8)
    local.writeUInt32LE(data.length, 18)
    local.writeUInt32LE(raw.length, 22)
    local.writeUInt16LE(nameBuf.length, 26)
    const central = Buffer.alloc(46)
    central.writeUInt32LE(0x02014b50, 0)
    central.writeUInt16LE(8, 10)
    central.writeUInt32LE(data.length, 20)
    central.writeUInt32LE(raw.length, 24)
    central.writeUInt16LE(nameBuf.length, 28)
    central.writeUInt32LE(offset, 42)
    locals.push(local, nameBuf, data)
    centrals.push(central, nameBuf)
    offset += 30 + nameBuf.length + data.length
  }
  const cd = Buffer.concat(centrals)
  const eocd = Buffer.alloc(22)
  eocd.writeUInt32LE(0x06054b50, 0)
  eocd.writeUInt16LE(Object.keys(files).length, 8)
  eocd.writeUInt16LE(Object.keys(files).length, 10)
  eocd.writeUInt32LE(cd.length, 12)
  eocd.writeUInt32LE(offset, 16)
  return Buffer.concat([...locals, cd, eocd])
}

describe('Office text', () => {
  test('docx: paragraphs, a table as tab-separated rows, entities and Vietnamese', () => {
    const doc = zip({
      'word/document.xml':
        '<w:document><w:body>' +
        '<w:p><w:r><w:t>Hợp đồng</w:t></w:r><w:r><w:t xml:space="preserve"> mua bán</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>A &amp; B &lt;x&gt;</w:t></w:r></w:p>' +
        '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Bên</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Vai trò</w:t></w:r></w:p></w:tc></w:tr>' +
        '<w:tr><w:tc><w:p><w:r><w:t>BIDV</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Bên vay</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
        '</w:body></w:document>',
    })
    const text = docxToText(readZip(doc))
    expect(text).toContain('Hợp đồng mua bán')
    expect(text).toContain('A & B <x>')
    expect(text).toContain('Bên\tVai trò')
    expect(text).toContain('BIDV\tBên vay')
  })

  test('xlsx: sheet names, shared and inline strings, numbers, and empty columns kept', () => {
    const book = zip({
      'xl/workbook.xml':
        '<workbook><sheets><sheet name="Nhân sự" sheetId="1" r:id="rId1"/><sheet name="Phí" sheetId="2" r:id="rId2"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels':
        '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="/xl/worksheets/sheet2.xml"/></Relationships>',
      'xl/sharedStrings.xml': '<sst><si><t>Tên</t></si><si><r><t>Phòng</t></r><r><t> ban</t></r></si><si><t>Minh</t></si></sst>',
      'xl/worksheets/sheet1.xml':
        '<worksheet><sheetData>' +
        '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row>' +
        '<row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>42</v></c><c r="C2" t="inlineStr"><is><t>IT</t></is></c></row>' +
        '</sheetData></worksheet>',
      'xl/worksheets/sheet2.xml': '<worksheet><sheetData><row r="1"><c r="A1"><v>100000</v></c></row></sheetData></worksheet>',
    })
    const text = xlsxToText(readZip(book))
    expect(text).toContain('## Nhân sự')
    expect(text).toContain('Tên\t\tPhòng ban')
    expect(text).toContain('Minh\t42\tIT')
    expect(text).toContain('## Phí\n100000')
  })

  test('pptx: slides in numeric order, not string order', () => {
    const deck = zip({
      'ppt/slides/slide10.xml': '<p:sld><a:p><a:r><a:t>Mười</a:t></a:r></a:p></p:sld>',
      'ppt/slides/slide2.xml': '<p:sld><a:p><a:r><a:t>Hai</a:t></a:r></a:p></p:sld>',
    })
    const text = pptxToText(readZip(deck))
    expect(text.indexOf('Slide 2')).toBeLessThan(text.indexOf('Slide 10'))
    expect(text).toContain('Hai')
  })

  test('a file that is not an archive, or the wrong kind, fails with a short reason', () => {
    expect(() => officeToText(Buffer.from('not a zip at all, just text'), 'docx')).toThrow(/not a zip/)
    expect(() => officeToText(zip({ 'x.txt': 'hi' }), 'docx')).toThrow(/not a Word document/)
  })

  test('only docx, xlsx and pptx are treated as Office files', () => {
    expect(isOfficeFile('C:\\inbox\\abc.DOCX')).toBe(true)
    expect(isOfficeFile('/inbox/abc.xlsx')).toBe(true)
    expect(isOfficeFile('/inbox/abc.pdf')).toBe(false)
    expect(isOfficeFile('/inbox/abc.doc')).toBe(false)
  })

  test('numeric entities decode', () => {
    expect(decodeXml('&#7871;&#x1EBF;')).toBe('ếế')
  })
})

describe('pasted images', () => {
  const PNG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0])
  const SERVICE = 'https://smba.trafficmanager.net/apac/'
  const IMAGE_URL = 'https://smba.trafficmanager.net/apac/v3/attachments/0-sa-abc/views/original'
  const activity = (attachments: unknown[], serviceUrl = SERVICE) => ({ id: '1793', serviceUrl, attachments })
  const token = async () => 'BOT_TOKEN'

  let inbox: string
  const realFetch = globalThis.fetch
  let calls: { url: string; auth: string | null }[]
  beforeEach(() => {
    inbox = mkdtempSync(join(tmpdir(), 'msteams-img-'))
    calls = []
  })
  afterEach(() => {
    globalThis.fetch = realFetch
  })
  const mockFetch = (res: () => Response) => {
    globalThis.fetch = (async (url: string, init?: RequestInit) => {
      calls.push({ url: String(url), auth: new Headers(init?.headers).get('authorization') })
      return res()
    }) as typeof fetch
  }

  test('an image/* attachment on the service host is registered as an image', () => {
    const files = new AttachmentHandles(undefined, token).register(
      activity([{ contentType: 'image/*', contentUrl: IMAGE_URL }, { contentType: 'text/html', content: '<img>' }]),
    )
    expect(files).toHaveLength(1)
    expect(files[0]!.kind).toBe('image')
  })

  test('an image on any other host, or with no bot token, is ignored', () => {
    const elsewhere = activity([{ contentType: 'image/png', contentUrl: 'https://evil.example.com/a.png' }])
    expect(new AttachmentHandles(undefined, token).register(elsewhere)).toHaveLength(0)
    const onService = activity([{ contentType: 'image/*', contentUrl: IMAGE_URL }])
    expect(new AttachmentHandles(undefined, undefined).register(onService)).toHaveLength(0)
  })

  test('download sends the bot token to the service host and names the file by its bytes', async () => {
    mockFetch(() => new Response(PNG, { status: 200 }))
    const handles = new AttachmentHandles(undefined, token)
    const [image] = handles.register(activity([{ contentType: 'image/*', contentUrl: IMAGE_URL }]))
    const path = await handles.download(image!.id, inbox)
    expect(path.endsWith('.png')).toBe(true)
    expect(readFileSync(path)).toEqual(PNG)
    expect(calls).toEqual([{ url: IMAGE_URL, auth: 'Bearer BOT_TOKEN' }])
  })

  test('a redirect is refused, so the token response never follows it', async () => {
    mockFetch(() => new Response(null, { status: 302, headers: { location: 'https://evil.example.com/x' } }))
    const handles = new AttachmentHandles(undefined, token)
    const [image] = handles.register(activity([{ contentType: 'image/*', contentUrl: IMAGE_URL }]))
    const err = await handles.download(image!.id, inbox).catch((e: Error) => e)
    expect((err as Error).message).toMatch(/redirected/)
    expect(calls).toHaveLength(1)
  })

  test('file attachments still never carry the bot token', async () => {
    mockFetch(() => new Response(PNG, { status: 200 }))
    const handles = new AttachmentHandles(['.sharepoint.com'], token)
    const [file] = handles.register({
      id: '1794',
      serviceUrl: SERVICE,
      attachments: [
        {
          contentType: 'application/vnd.microsoft.teams.file.download.info',
          name: 'a.png',
          content: { downloadUrl: 'https://contoso-my.sharepoint.com/x?tempauth=t', fileType: 'png', uniqueId: 'u1' },
        },
      ],
    })
    await handles.download(file!.id, inbox)
    expect(calls[0]!.auth).toBeNull()
  })

  test('sniffImageExt recognises the common formats', () => {
    expect(sniffImageExt(PNG)).toBe('png')
    expect(sniffImageExt(Buffer.from([0xff, 0xd8, 0xff, 0xe0]))).toBe('jpg')
    expect(sniffImageExt(Buffer.from('GIF89a'))).toBe('gif')
    expect(sniffImageExt(Buffer.from('hello'))).toBeUndefined()
  })
})
