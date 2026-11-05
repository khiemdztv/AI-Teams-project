/**
 * Colleagues by name (fork addition, 05/11), shared by Teams messages, email
 * and meetings: "Jimmy", "kien", "Hà Nguyễn" become one person in the firm's
 * directory, or an error the sender has to answer.
 *
 * A name is searched in display names and email addresses. An exact display
 * name (accents ignored) wins over partial matches; several candidates are
 * never narrowed down here, they are listed so the sender picks. Guest
 * accounts are in the directory too and are left out, so a name only ever
 * finds a colleague: someone outside the firm is reached by email address.
 */

import { GraphError, parseRecipient } from './mail.js'

/** Candidates listed when a name matches several people. */
export const MAX_MATCHES = 8

export type Person = { id: string; name: string; email: string }

/** A GET against Graph v1.0 with the acting person's token. */
export type GraphGet = (path: string, headers?: Record<string, string>) => Promise<any>

export function toPerson(u: any): Person {
  return { id: String(u.id), name: String(u.displayName ?? ''), email: String(u.mail || u.userPrincipalName || '').toLowerCase() }
}

/** Lower case without Vietnamese accents, for comparing names. */
export function foldName(s: string): string {
  return s.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/gi, 'd').toLowerCase().trim()
}

/** An entry with an "@" is an address ("a@b.com", "Name <a@b.com>"); anything else is a name. */
export function isAddress(entry: string): boolean {
  return String(entry ?? '').includes('@')
}

export class People {
  constructor(
    private readonly get: GraphGet,
    private readonly internalDomains: string[],
  ) {}

  /** Addresses outside the firm's domains. */
  external(list: string[]): string[] {
    const internal = this.internalDomains.map(d => d.toLowerCase())
    return list.filter(a => a.includes('@') && !internal.includes(a.split('@')[1] ?? ''))
  }

  /** The directory entry for an address, or null when the directory has none. */
  async byEmail(address: string): Promise<Person | null> {
    const email = parseRecipient(address).emailAddress.address.toLowerCase()
    try {
      return toPerson(await this.get(`/users/${encodeURIComponent(email)}?$select=id,displayName,mail,userPrincipalName`))
    } catch (err) {
      if (err instanceof GraphError && err.status === 404) return null
      throw err
    }
  }

  /** The one colleague a name points at; otherwise an error listing the candidates. */
  async byName(entry: string): Promise<Person> {
    const q = String(entry ?? '').trim().replace(/["\\]/g, '')
    if (!q) throw new Error('an empty name')
    const search = encodeURIComponent(`"displayName:${q}" OR "mail:${q}"`)
    const res = await this.get(`/users?$search=${search}&$select=id,displayName,mail,userPrincipalName&$top=${MAX_MATCHES + 1}`, {
      ConsistencyLevel: 'eventual',
    })
    const found = ((res?.value ?? []) as any[]).map(toPerson).filter(u => u.email && !this.external([u.email]).length)
    const exact = found.filter(u => foldName(u.name) === foldName(q))
    const pick = exact.length === 1 ? exact : found
    if (pick.length === 1) return pick[0]!
    if (!pick.length) throw new Error(`nobody called "${q}" in the firm's directory; ask the sender for the email address`)
    throw new Error(
      `"${q}" matches several people; ask the sender which one and pass their email: ${pick
        .slice(0, MAX_MATCHES)
        .map(u => `${u.name} <${u.email}>`)
        .join(', ')}${pick.length > MAX_MATCHES ? ', …' : ''}`,
    )
  }

  /**
   * An email or meeting recipient: an address is kept exactly as given (people
   * outside the firm included), a name becomes the colleague it points at.
   */
  async recipient(entry: string): Promise<{ address: string; name?: string }> {
    if (isAddress(entry)) {
      const a = parseRecipient(entry).emailAddress
      return a.name ? { address: a.address, name: a.name } : { address: a.address }
    }
    const p = await this.byName(entry)
    return p.name ? { address: p.email, name: p.name } : { address: p.email }
  }

  async recipients(entries: string[] | undefined): Promise<{ address: string; name?: string }[]> {
    const out: { address: string; name?: string }[] = []
    for (const e of entries ?? []) out.push(await this.recipient(e))
    return out
  }

  /** Lower-case addresses, sorted: what expect_* lists are compared as. */
  async addresses(entries: string[] | undefined): Promise<string[]> {
    return (await this.recipients(entries)).map(r => r.address.toLowerCase()).sort()
  }
}
