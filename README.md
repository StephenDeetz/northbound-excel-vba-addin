# Northbound Excel Add-in

A public Excel add-in from Northbound Group. This repo holds the add-in's
VBA source (`.bas`/`.cls` modules and Ribbon XML), version-controlled
separately from the built `.xlam` file.

## Features

### Pretty Print / Minify

Reformats Excel formulas for readability (Pretty Print), or collapses them
back down (Minify). Available from the Northbound ribbon tab, at the cell,
range, sheet, and workbook level.

Pretty Print makes complex formulas easier to read.

| Before | After |
|---|---|
| <img src="images/pretty-print-1-before.png" height="150"> | <img src="images/pretty-print-1-after.png" height="150"> |
| <img src="images/pretty-print-2-before.png" height="150"> | <img src="images/pretty-print-2-after.png" height="200"> |
| <img src="images/pretty-print-3-before.png" height="150"> | <img src="images/pretty-print-3-after.png" height="300"> |
| <img src="images/pretty-print-4-before.png" height="150"> | <img src="images/pretty-print-4-after.png" height="150"> |

More features will be added here as they are carved out of Northbound's
internal add-in.

## Installation

1. Download the latest release:
   https://github.com/StephenDeetz/northbound-excel-vba-addin/releases/latest
2. Verify the download's SHA-256 hash matches the value published on the
   release page. On Windows: `certutil -hashfile Northbound.xlam SHA256`
3. Add the `.xlam` file as an Excel add-in. For step-by-step instructions,
   including unblocking the file and setting up a trusted location, see
   [INSTALL.md](INSTALL.md).

## Source Code

This repo contains the add-in's source for transparency and version
history. There is no build step -- the `.xlam` is assembled in Excel's
VBA editor (VBE) and its modules are exported here as plain text files.
See `CLAUDE.md` for the development workflow.

## License

MIT -- see [LICENSE](LICENSE).

## About Northbound

Northbound Group advises founders of eCommerce, SaaS, and tech-enabled businesses on exits,
acquisitions, and strategic decisions that define long-term value.

As a financial services company, Northbound works heavily in Excel. We come across Excel
issues of all types - from slow workbooks to complex formulas. Because of this, Northbound
is happy to share this tool with our fellow Excel travelers.

## Links

- Website: [https://northboundgroup.com/](https://northboundgroup.com/?utm_source=github&utm_medium=readme&utm_campaign=vba_addin)
- Issues / Source: https://github.com/StephenDeetz/northbound-excel-vba-addin
