# TermStoreTools

Console tool for SharePoint Online Term Store maintenance - full **Create / Rename / Update /
Delete** coverage for Term Groups, Term Sets, and Terms, plus Move for Term Sets and Terms. Uses
`PnP.Framework` for interactive auth (replaces the older, deprecated
`OfficeDevPnP.Core.AuthenticationManager`); Taxonomy operations themselves are still plain CSOM
(`Microsoft.SharePoint.Client.Taxonomy`), same as before - PnP.Framework only replaces how the
tool signs in.

## Run it

```
TermStoreTools.exe
```

You'll be prompted for a site URL (interactive/web login via the well-known "PnP Management
Shell" multi-tenant Azure AD app - no app registration of your own needed), then a menu:

| | Create | Rename | Update | Delete | Move |
|---|---|---|---|---|---|
| **Term Group** | 1 | 2 | 3 (description) | 4 | - |
| **Term Set** | 5 | 6 | 7 (description) | 8 | 9 (to another Group) |
| **Term** | 10 | 11 (adds a label) | 12 (description), 13 (deprecate), 14 (reactivate) | 15 | 16 (to another Term Set) |

Plus **17) List / Export Terms** in a Term Set (prints to console; optionally exports to CSV).

Each option prompts for the term store / group / term set / term names it needs.

## Term Id shortcut

Any operation on a specific **Term** (11-16) asks:
```
Enter Term Id (GUID), or leave blank to browse by Group/TermSet/Term name =>
```
Paste a Term's GUID and it resolves the Term directly via `TermStore.GetTerm(id)` - no need to
know which Group/TermSet it lives in. Leave it blank to fall back to the usual
Group -> TermSet -> Term name walk.

## Caching

The Term Store, and every Group/TermSet/Term you look up by name or Id, are cached for the rest
of the run - looking the same one up again (e.g. adding several terms to the same TermSet, or
running List after Create) reuses the cached object instead of re-prompting for the TermStore
name or re-fetching from the server. The cache is per-run only (in-memory, cleared on exit); a
rename/move/delete updates or evicts the affected cache entries so stale data isn't reused.

## Notes

- Connects via `PnP.Framework.AuthenticationManager.CreateWithInteractiveLogin` using
  `AuthenticationManager.CLIENTID_PNPMANAGEMENTSHELL` - a public, Microsoft-approved multi-tenant
  app ID used throughout the PnP community specifically for this kind of interactive tool; you
  don't need to register your own Azure AD app.
- Delete operations (4, 8, 15) are destructive and permanent (`DeleteObject()` +
  `ExecuteQuery()`) - each requires typing `yes` to proceed, but there is no undo once confirmed.
  Deleting a Group or Term Set also deletes everything inside it.
- Renaming a **Term** (11) adds a **new label**; it doesn't remove or overwrite the term's
  existing labels unless the new one is marked default and used in the same locale. Renaming a
  **Term Set or Term Group** (2/6) is different - both just overwrite the single `Name` property
  directly (no multi-label system at that level), same for their descriptions (3/7 - plain
  `Description` property vs. Term's locale-aware `SetDescription(text, lcid)`, option 12).
- Deprecating a term (13) hides it from tagging but keeps it and its history intact - prefer
  this over deleting when you're not sure a term is still in use.

## Status
Full Create/Rename/Update/Delete matrix implemented for Term Group, Term Set, and Term; project
builds cleanly (`dotnet build`, 0 errors) against the PnP.Framework-based reference set. Not yet
exercised end-to-end against a live tenant in this pass - test the create/delete paths against a
non-production term store first.
