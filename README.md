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

### Git and Excel

### Release Workflow

1. Utilize `git-flow` to create *release/* branches off the default *develop* branch
2. Close release branches adding a corresponding `git tag` for the version number.
3. Push to remote
4. Run command `gh release create <version> --generate-notes --title 'Version <version>'` from terminal using *GitHub-CLI*.
5. This will invoke the [release-xl.yml](.github/workflows/release-xl.yml) GitHub Action workflow to add the versioned workbook to the release assets.


