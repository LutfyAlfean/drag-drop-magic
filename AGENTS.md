<!-- LOVABLE:BEGIN -->
> [!IMPORTANT]
> This project is connected to [Lovable](https://lovable.dev). Avoid rewriting
> published git history — force pushing, or rebasing/amending/squashing commits
> that are already pushed — as it rewrites history on Lovable's side and the
> user will likely lose their project history.
>
> Commits you push to the connected branch sync back to Lovable and show up in
> the editor, so keep the branch in a working state.
<!-- LOVABLE:END -->

- Keep spreadsheet parsing and LDIF generation in the browser; uploaded directory data must not be sent to a server.
- Use the shared LDIF module for validation and serialization so preview and downloaded output remain identical.
- Export inetOrgPerson by default, not posixAccount, because the input lacks required POSIX numeric IDs and home directories.
