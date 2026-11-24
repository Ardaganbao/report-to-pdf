# Action-tracker in a report (`mose-actions`)

A Markdown extension for the "Suivi des actions" view: a Red/Orange/Green table of issues and
actions, like the `Suivi_actions_*.xlsx` follow-up sheet, with a **Synthèse** line of the counts
underneath. It renders both on the web report (gi.grassone.be) and in the printed PDF.

## Writing one

In the report, add a fenced block with the language `mose-actions`:

````markdown
```mose-actions
# Suivi des actions — sur base du rapport du 09/09/2026
Issue | Solution | Owner | Due date | Status | Comments
Token de téléchargement à fournir | Fournir le token | Quentin Buis |  | Orange | demande explicite, pas de clôture
Mises à jour vCenter/Windows/Veeam | Exécuter et confirmer | ONIT |  | Green | reprises dans les actions du jour
Garanties matériel expirées | Vérifier et régulariser |  | 27/06/2026 | Red | extension fournie, enregistrement nécessaire
```
````

Rules:

- **One row per line**, columns separated by `|`. Leading/trailing `|` are optional, so a normal
  Markdown table pasted in works too (a `---|---` divider line is ignored).
- A line starting with `#` is the **table title** (optional).
- The **first line that names the columns** is the header. If you leave it out, the columns are
  assumed to be `Issue | Solution | Owner | Due date | Status | Comments` (the workbook's order).
- Any number of columns works. The **Status** column is the one whose header contains `status` or
  `statut` (or, with no header, the 5th column); if none is named, the column whose values all read
  as a status is used.
- Empty cells are allowed (leave the space between the `|` blank).

## Status values

Case-insensitive. The row is tinted and a coloured pill is shown.

| Colour | Accepted values |
|---|---|
| 🔴 Red | `red`, `rouge`, `R`, `critical`, `high`, `urgent`, `ko`, `nok` |
| 🟠 Orange | `orange`, `amber`, `O`, `A`, `medium`, `warning`, `attention`, `en cours`, `ongoing`, `wip` |
| 🟢 Green | `green`, `vert`, `G`, `V`, `ok`, `done`, `résolu`, `clos`, `fait`, `low` |

Anything else is shown uncoloured. The pill keeps the word you typed (e.g. `Orange`).

## Synthèse

A line is added below the table with: **Total**, the count per colour, how many rows have **no
owner**, and how many have a **due date**. The owner/due counts appear only when those columns exist.

## Where it is rendered

- **PDF** — [htmltoPDF.ps1](htmltoPDF.ps1) (`Expand-MoseActions` / `ConvertTo-MoseActionsTable`)
  builds the table server-side from the fetched HTML, so wkhtmltopdf gets a plain styled table (no
  JavaScript, same approach as the Mermaid/graph handling).
- **Web** — `mose-actions.js`, in the **mose-wordpress** plugin, builds the same table in the
  browser. The plugin enqueues it in `my_admin_scripts()` (`multi-block.php`), right after `view.js`,
  on report pages only (`ardareport_is_report_page()`). It is standalone — no Chart.js/Mermaid
  dependency.

The two builders are kept in step: if you change the columns, colours or Synthèse in one, change the
other to match.

## Wide tables

Six columns of prose are wide. Wrap the block between the report's landscape tokens if it overflows
in portrait:

```
<!--LANDSCAPE-->
...the mose-actions block...
<!--PORTRAIT-->
```
