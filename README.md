<h1 align="center"> VBASuite </h1>

<p align="center"> 
  <img src="https://th.bing.com/th/id/OIP.ZP7URHglG8ZknZorYplFZwHaER?pid=ImgDet&rs=1" height="50%" width="50%" />
</p>

<hr>

## Contents

## Overview

## Features

## Development

### VBA Modules

- [modUtils](src/VBASuite.xlsm/modUtils.bas): Utilities Module
- [modSetup](src/VBASuite.xlsm/modSetup.bas): Module housing various setup workflows
- [modOptimize](src/VBASuite.xlsm/modOptimize.bas): VBA optimization module

### Git and Excel

This repository utilizes the following resources for optimal VBA, Excel, and Git Integrations:

- [Git Flow]()
- [Git LFS]()
- [GitHub CLI]()
- [Git XL](https://www.xltrail.com/git-xl)
- [vbaDeveloper Excel AddIn]()
- [xvba VSCode Addins]()

Other tools worth mentioning:

- [ImportExcel PowerShell Module]()
- [vba-blocks CLI tool]()

#### Setup

1. Download and install [Git XL](https://www.xltrail.com/git-xl) from the [xltrail](https://www.xltrail.com/) website.
2. Initialize Git XL via `git xl install` and then `git xl install --local`.
3. Initialize Git-LFS and Git-Flow on the repository, add respective `.gitattributes`.

### Release Workflow

1. Utilize `git-flow` to create *release/* branches off the default *develop* branch
2. Close release branches adding a corresponding `git tag` for the version number.
3. Push to remote
4. Run command `gh release create <version> --generate-notes --title 'Version <version>'` from terminal using *GitHub-CLI*.
5. This will invoke the [release-xl.yml](.github/workflows/release-xl.yml) GitHub Action workflow to add the versioned workbook to the release assets.


