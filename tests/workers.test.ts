/** Fork (20/11): workers.json, the per-person session table (src/workers.ts). */

import { describe, expect, test } from 'bun:test'
import { mkdtempSync, writeFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { ownsJob, parseWorkers, readWorkers, workerOf } from '../src/workers.js'

const A = '11111111-1111-1111-1111-111111111111'
const B = '22222222-2222-2222-2222-222222222222'

describe('parseWorkers', () => {
  test('keeps valid entries, lowercases owners, defaults the shared port', () => {
    const { workers, errors } = parseWorkers({ workers: [{ code: 'an', owner: A.toUpperCase(), port: 4001 }] })
    expect(errors).toEqual([])
    expect(workers).toEqual({ defaultPort: 4000, workers: [{ code: 'an', owner: A, port: 4001 }] })
  })

  test('drops malformed and duplicate entries, with a reason each', () => {
    const { workers, errors } = parseWorkers({
      defaultPort: 4000,
      workers: [
        { code: 'An', owner: A, port: 4001 },
        { code: 'an', owner: 'not-an-id', port: 4001 },
        { code: 'an', owner: A, port: 4000 },
        { code: 'an', owner: A, port: 4001 },
        { code: 'an', owner: B, port: 4002 },
        { code: 'binh', owner: A, port: 4003 },
      ],
    })
    expect(workers.workers.map(w => w.code)).toEqual(['an'])
    expect(errors).toHaveLength(5)
  })
})

describe('readWorkers', () => {
  test('no file means no workers and no error; a broken file is reported', () => {
    const dir = mkdtempSync(join(tmpdir(), 'workers-'))
    expect(readWorkers(dir)).toEqual({ workers: { defaultPort: 4000, workers: [] }, errors: [] })
    writeFileSync(join(dir, 'workers.json'), '{nope')
    expect(readWorkers(dir).errors).toHaveLength(1)
    writeFileSync(join(dir, 'workers.json'), JSON.stringify({ defaultPort: 4100, workers: [{ code: 'an', owner: A, port: 4101 }] }))
    expect(readWorkers(dir).workers.workers).toHaveLength(1)
  })
})

describe('routing helpers', () => {
  const { workers } = parseWorkers({ workers: [{ code: 'an', owner: A, port: 4001 }] })

  test('workerOf matches the owner whatever the case', () => {
    expect(workerOf(workers, A.toUpperCase())?.code).toBe('an')
    expect(workerOf(workers, B)).toBeUndefined()
    expect(workerOf(workers, undefined)).toBeUndefined()
  })

  test("a person's session runs only their jobs; the shared one runs everyone else's", () => {
    expect(ownsJob(workers, A, A)).toBe(true)
    expect(ownsJob(workers, A, B)).toBe(false)
    expect(ownsJob(workers, undefined, A)).toBe(false)
    expect(ownsJob(workers, undefined, B)).toBe(true)
  })
})
