import { describe, expect, it } from 'vitest';
import { escapeDn, generateLdif, ldifLine, validBaseDn } from './ldif';
const settings = { baseDn: 'dc=sekolah,dc=id', ou: 'users', passwordMode: 'suffix', customPassword: '' };
const row = { username: 'ahmad', firstname: 'Ahmad', lastname: 'Fauzi', email: 'ahmad@example.com' };
describe('LDIF conversion', () => {
 it('exports inetOrgPerson and default password without invalid posixAccount', () => { const value = generateLdif([row], settings); expect(value).toContain('dn: uid=ahmad,ou=users,dc=sekolah,dc=id'); expect(value).toContain('userPassword: ahmad123'); expect(value).not.toContain('posixAccount'); });
 it('skips incomplete rows and duplicate usernames', () => { const value = generateLdif([row, row, {...row, lastname: ''}], settings); expect(value.match(/dn:/g)).toHaveLength(1); });
 it('supports no OU and custom password', () => { const value = generateLdif([row], {...settings, ou:'', passwordMode:'custom', customPassword:'secret'}); expect(value).toContain('dn: uid=ahmad,dc=sekolah,dc=id'); expect(value).toContain('userPassword: secret'); });
 it('escapes DN metacharacters and encodes unicode or injected newlines', () => { expect(escapeDn(' a,b ')).toBe('\\ a\\,b\\ '); expect(ldifLine('cn', 'José')).toBe('cn:: Sm9zw6k='); expect(ldifLine('cn', 'a\nb')).toBe('cn:: YQpi'); });
 it('can omit passwords', () => { expect(generateLdif([row], {...settings, passwordMode:'none'})).not.toContain('userPassword'); });
 it('validates the base DN', () => { expect(validBaseDn('dc=sekolah,dc=ac,dc=id')).toBe(true); expect(validBaseDn('')).toBe(false); expect(validBaseDn('sekolah')).toBe(false); });
});
