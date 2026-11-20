# Modern Content Editor Web Part

## Summary

This webpart replicates the classic SharePoint Content Editor Web part and lets you link to a HTML file containing HTML / JS or insert HTML code directly into the page.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.21.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

Version 1.0.0.2 is built with an SPfx version that only works on SPO. Previous version should work with SharePoint 2019 on-prem.

## Content Security Policy (CSP)

SharePoint Online enforces a Content Security Policy from March 1, 2026. This affects the code you add to the web part:

- **External scripts** (for example from a CDN or another tenant) only load if the domain is added to **Trusted script sources** in the SharePoint admin center (Advanced > Script sources). Scripts stored in your own tenant work without extra setup.
- **Inline `<script>` blocks** are supported. The web part runs them in the global scope, so top-level functions and variables work like in the classic CEWP.
- **Inline event handlers are blocked**, for example `<button onclick="doSomething()">`, and so are `href="javascript:..."` links. Attach the handlers from a script instead:
  ```html
  <button id="myButton">Click me</button>
  <script>
    document.getElementById("myButton").addEventListener("click", doSomething);
  </script>
  ```
- Only JavaScript `<script>` blocks are run. Blocks with another `type` (for example `text/html` templates or `application/json`) are left in place for your code to read. Inline `<script type="module">` is not supported, so use an external file with `src` instead.

The version in use is shown in the web part when the page is in edit mode, and at the top of the property pane.

## Solution

Modern CEWP | Alexander Bautz - SPJSWorks.com

## Version history
1.0.0.7 - November 19, 2026
Fixed loading of HTML files that reference external stylesheets with `<link rel="stylesheet">`. Relative stylesheet and script URLs are now resolved against the linked HTML file.
Inline scripts now run in the global scope again, so functions and variables declared in one script can be used from other scripts.
Script blocks that are not JavaScript (templates, JSON) are no longer run.
Fixed the "empty" message not showing on a newly added web part, and the error message showing "undefined" when a content link failed to load.
The web part version is now shown in edit mode and in the property pane.
A script that fails to load (for example blocked by CSP) is logged to the console, and the remaining scripts still run.
jQuery is now also available as `$` (only if `$` is not already defined).
The web part now renders in the local workbench.

1.0.0.6 - February 27, 2026
Fixed a bug where the <style> tags were not added to the DOM when loaded as part as a HTML file.

1.0.0.5 - February 22, 2026
Made it compatible with the new CSP policy that Microsoft starts enforcing on March 1, 2026. Added title property on the webpart that can be used to describe what the code does for easy maintenance on a page with multiple webparts.

1.0.0.2 - August 08, 2023
Updated SPfx version and rebuild solution. Published to GitHub

1.1 -  March 10, 2021
bugfix 

1.0 - January 29, 2021
Initial release

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## License
Copyright 2021-2026 - Alexander Bautz / SPJSWorks.com

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **npm run serve** (uses spfx-fast-serve)
- To build a package for deployment:
  - **gulp bundle --ship**
  - **gulp package-solution --ship**
  - The package is created at `sharepoint/solution/modern-cewp.sppkg`

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
