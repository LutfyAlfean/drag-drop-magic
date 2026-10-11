export type DataRow = Record<string, string>;
export const requiredColumns = ['username', 'firstname', 'lastname', 'email'];
export type LdifSettings = { baseDn: string; ou: string; passwordMode: string; customPassword: string };
export function escapeDn(value: string) {
  return value.replace(/[,+=<>;"\\]/g, '\\$&').replace(/^ /, '\\ ').replace(/^#/, '\\#').replace(/ $/, '\\ ').replace(/\0/g, '\\00');
}
export function ldifLine(attribute: string, value: string) {
  if (/^[ :<]|[\r\n\0]|[^\x20-\x7e]| $/.test(value)) {
    const bytes = new TextEncoder().encode(value);
    return `${attribute}:: ${btoa(Array.from(bytes, b => String.fromCharCode(b)).join(''))}`;
  }
  return `${attribute}: ${value}`;
}
export function isValidRow(row: DataRow) { return Boolean(row.username?.trim() && row.lastname?.trim()); }
export function generateLdif(rows: DataRow[], settings: LdifSettings) {
  const seen = new Set<string>();
  return rows.filter(isValidRow).filter(row => { const uid = row.username.trim().toLowerCase(); if (seen.has(uid)) return false; seen.add(uid); return true; }).map(row => {
    const username = row.username.trim();
    const parent = settings.ou.trim() ? `ou=${escapeDn(settings.ou.trim())},${settings.baseDn.trim()}` : settings.baseDn.trim();
    const entries = [ldifLine('dn', `uid=${escapeDn(username)},${parent}`), 'objectClass: inetOrgPerson', ldifLine('cn', `${row.firstname?.trim() || ''} ${row.lastname.trim()}`.trim()), ldifLine('sn', row.lastname.trim()), ldifLine('uid', username)];
    if (row.email?.trim()) entries.push(ldifLine('mail', row.email.trim()));
    const password = settings.passwordMode === 'suffix' ? username + '123' : settings.passwordMode === 'username' ? username : settings.customPassword;
    if (settings.passwordMode !== 'none') entries.push(ldifLine('userPassword', password));
    return entries.join('\n');
  }).join('\n\n') + '\n\n';
}
export function validBaseDn(value: string) { return /^(?:[a-z][a-z0-9-]*=(?:\\.|[^,=\r\n])+)(?:,[a-z][a-z0-9-]*=(?:\\.|[^,=\r\n])+)*$/i.test(value.trim()); }
