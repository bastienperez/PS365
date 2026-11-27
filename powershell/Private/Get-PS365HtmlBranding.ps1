function Get-PS365HtmlBranding {
    <#
    .SYNOPSIS
    Returns the shared PS365 design of the HTML reports: base CSS (365Ops look), the top bar brand block
    (PS365 logo, 'by Clidsys'), the page heading and the footer.

    .DESCRIPTION
    Every interactive HTML report of the module (Get-SPOSiteReport, Get-SPOStorageReport, Get-NestedGroup)
    builds its page from these pieces so they share one look, taken from the 365Ops design: a quiet white top
    bar with the PS365 logo and the 'by Clidsys' link, a large page title, white cards with 12px corners on a
    light grey page, Plus Jakarta Sans, status dots and pills, and the Clidsys orange for the primary button
    and the links. Each page keeps its own layout-specific CSS on top.

    The logo and the font are embedded as base64 data URIs (Private/Assets): the page stays a single file
    that works offline. Plus Jakarta Sans is under the SIL Open Font License 1.1
    (Private/Assets/fonts/PlusJakartaSans-LICENSE.txt). As in 365Ops, each weight is drawn one step heavier
    (400 uses the 500 file, 600 the 700 file, 700 the 800 file): the font reads thin otherwise.

    Returned object:
    - Css: font faces, tokens (:root variables) and base styles. Pages reference var(--surface), var(--border)...
    - HeadHtml: the <link rel="icon"> tag (PS365 logo as the browser tab icon), to place inside <head>.
    - HeaderBrandHtml: the left part of the top bar (logo, PS365 by Clidsys), inside <header class="ps365-header">.
    - PageHeadHtml: the page heading (title, subtitle, optional action buttons).
    - FooterHtml: the page footer (generation date, module version, links).
    - CopyButtonHtml: a 'Copy' button; it copies the closest element carrying the data-copy-root attribute.
    - ScriptHtml: the copy logic (ps365CopyBlock), to place once before </body> on a page using copy buttons.

    .PARAMETER Title
    Report title, shown as the page heading, for example 'SharePoint Site Report'.

    .PARAMETER Subtitle
    (Optional) Line under the title, plain text (HTML-encoded here), for example the generation date and scope.

    .PARAMETER ActionsHtml
    (Optional) HTML of the buttons shown on the right of the page heading. Inserted as-is: callers only pass
    their own static markup.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $false)]
        [string]$Subtitle,

        [Parameter(Mandatory = $false)]
        [string]$ActionsHtml
    )

    $ps365Url = 'https://ps365.clidsys.com'
    $clidsysUrl = 'https://clidsys.com'
    $assetsPath = Join-Path -Path $PSScriptRoot -ChildPath 'Assets'

    $logoDataUri = ''
    $logoPath = Join-Path -Path $assetsPath -ChildPath 'PS365-logo.svg'
    if (Test-Path -Path $logoPath -PathType Leaf) {
        $logoDataUri = 'data:image/svg+xml;base64,' + [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($logoPath))
    }

    # CSS weight -> font file, one step heavier as in 365Ops
    $fontFaces = [System.Text.StringBuilder]::new()
    foreach ($fontWeight in @(@{ Css = 400; File = 500 }, @{ Css = 600; File = 700 }, @{ Css = 700; File = 800 })) {
        $fontPath = Join-Path -Path $assetsPath -ChildPath "fonts\plus-jakarta-sans-latin-$($fontWeight.File)-normal.woff2"
        if (Test-Path -Path $fontPath -PathType Leaf) {
            $fontBase64 = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($fontPath))
            $null = $fontFaces.Append("    @font-face { font-family: 'Plus Jakarta Sans'; font-style: normal; font-display: swap; font-weight: $($fontWeight.Css); src: url(data:font/woff2;base64,$fontBase64) format('woff2'); }`n")
        }
    }

    $moduleVersion = (Get-Module -Name PS365 | Select-Object -First 1).Version
    $versionText = if ($moduleVersion) { " v$moduleVersion" } else { '' }
    $generatedOn = Get-Date -Format 'yyyy-MM-dd HH:mm'

    # Tokens from the 365Ops design (docs/design of 365ops): neutral greys, Clidsys orange #de761d as the brand
    # accent, its darker step #B85C10 for white-on-orange fills (3.1:1 otherwise) and #9A4A0A for links.
    # The role colors (admins, owners...) are solid fills carrying white text.
    $css = $fontFaces.ToString() + @'
    :root {
      --font: 'Plus Jakarta Sans', 'Segoe UI Variable Text', 'Segoe UI', -apple-system, system-ui, sans-serif;
      --mono: Consolas, 'Cascadia Mono', Menlo, monospace;
      --bg: #F4F5F7; --surface: #ffffff; --surface2: #FAFAFB; --hover: #F4F5F7;
      --border: #E3E5EA; --border-strong: #D9DCE2; --divider: #ECEEF1;
      --text: #16181D; --muted: #4B5160; --sub: #5B6270;
      --brand: #de761d; --brand-solid: #B85C10; --brand-hover: #9A4A0A; --accent: #9A4A0A; --accent-light: #FEF3E2;
      --flag: #C2410C; --ok: #1A7F4B; --skip: #C9CCD3;
      --radius: 8px; --card-radius: 12px; --shadow: none;
      --admins: #B91C1C; --owners: #B45309; --members: #15803D; --visitors: #4B5160;
    }
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: var(--font); background: var(--bg); color: var(--text); font-size: 14px; line-height: 1.45; }
    a { color: var(--accent); }
    a:hover { color: #6E3407; }
    :focus-visible { outline: 2px solid var(--brand-solid); outline-offset: 2px; }

    /* Top bar: white and quiet, the brand on the left, the page tools on the right */
    header.ps365-header { background: var(--surface); border-bottom: 1px solid var(--border); padding: 10px 24px; min-height: 56px; display: flex; align-items: center; gap: 12px; flex-wrap: wrap; position: sticky; top: 0; z-index: 10; }
    .ps365-brand { display: flex; align-items: center; gap: 10px; flex-shrink: 0; text-decoration: none; color: var(--text); }
    .ps365-brand img { width: 30px; height: 30px; display: block; }
    .ps365-brand-name { font-size: 17px; font-weight: 700; letter-spacing: -0.01em; color: var(--text); }
    .ps365-brand-by { font-size: 12.5px; color: var(--sub); }
    .ps365-brand-by a { color: var(--brand-solid); font-weight: 600; text-decoration: none; }
    .ps365-brand-by a:hover { text-decoration: underline; text-underline-offset: 2px; }
    header.ps365-header .spacer { flex: 1; }
    header.ps365-header .sep { width: 1px; height: 24px; background: var(--border); flex-shrink: 0; }

    /* Buttons, chips, pills, status dots (365Ops) */
    .btn { display: inline-flex; align-items: center; gap: 8px; height: 36px; padding: 0 14px; border-radius: var(--radius); border: 1px solid var(--border-strong); background: var(--surface); color: var(--text); font: inherit; font-weight: 600; cursor: pointer; white-space: nowrap; }
    .btn:hover { background: var(--hover); }
    .btn-primary { background: var(--brand-solid); border-color: var(--brand-solid); color: #fff; }
    .btn-primary:hover { background: var(--brand-hover); }
    .chip { display: inline-flex; align-items: center; height: 26px; padding: 0 10px; border-radius: 13px; border: 1px solid var(--border); background: var(--surface2); color: var(--muted); font-size: 12.5px; font-weight: 600; white-space: nowrap; flex-shrink: 0; }
    .chip.muted { font-weight: 400; }
    .pill { display: inline-block; padding: 2px 8px; border-radius: 5px; font-size: 12px; font-weight: 600; white-space: nowrap; }
    .pill-crit { background: #16181D; color: #fff; }
    .pill-high { background: #FDECEA; color: #9B1C1C; }
    .pill-med { background: #FEF3E2; color: #92400E; }
    .pill-low { background: #EEF0F3; color: #3D4350; }
    .st { display: inline-flex; align-items: center; gap: 6px; font-weight: 600; white-space: nowrap; }
    .dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; display: inline-block; }
    input[type=text], input[type=search] { height: 36px; padding: 0 12px; border: 1px solid var(--border-strong); border-radius: var(--radius); background: var(--surface); font: inherit; color: var(--text); outline: none; min-width: 0; }
    input[type=text]:focus, input[type=search]:focus { border-color: var(--brand-solid); box-shadow: 0 0 0 3px rgba(222,118,29,.15); }

    /* Page layout: large title, white cards on the grey page */
    .ps365-page { max-width: 1180px; margin: 0 auto; padding: 28px 24px 8px; }
    .page-head { display: flex; flex-wrap: wrap; align-items: flex-end; justify-content: space-between; gap: 16px; margin-bottom: 20px; }
    .page-head h1 { font-size: 28px; font-weight: 700; letter-spacing: -0.02em; line-height: 1.2; }
    .page-sub { margin-top: 6px; color: var(--muted); }
    .page-actions { display: flex; gap: 8px; flex-wrap: wrap; }
    .card { background: var(--surface); border: 1px solid var(--border); border-radius: var(--card-radius); }
    .card-toolbar { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; padding: 12px 16px; border-bottom: 1px solid var(--border); }
    section { margin-bottom: 24px; }
    section > h2 { font-size: 18px; font-weight: 700; letter-spacing: -0.01em; margin-bottom: 12px; }
    .section-head { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 12px; }
    .section-head h2 { font-size: 18px; font-weight: 700; letter-spacing: -0.01em; }

    /* Copy buttons: copy a block as rich HTML (pastes as real tables in Word) and as plain text */
    .btn-copy { height: 30px; padding: 0 10px; font-size: 12.5px; gap: 6px; color: var(--muted); }
    .btn-copy svg { width: 14px; height: 14px; }
    .btn-copy.done { color: var(--ok); border-color: var(--ok); }

    /* Tables (365Ops) */
    table.ps365-table { width: 100%; border-collapse: collapse; }
    .ps365-table th { text-align: left; font-size: 12px; font-weight: 600; color: var(--muted); padding: 10px 14px; border-bottom: 1px solid var(--border); background: var(--surface2); }
    .ps365-table td { padding: 12px 14px; border-bottom: 1px solid var(--divider); vertical-align: top; }
    .ps365-table tr:last-child td { border-bottom: none; }
    .ps365-table tbody tr:hover td { background: var(--surface2); }
    .ps365-table .id { font-size: 12px; color: var(--sub); }

    footer.ps365-footer { max-width: 1180px; margin: 8px auto 28px; padding: 16px 24px 0; display: flex; align-items: center; gap: 10px; flex-wrap: wrap; font-size: 12.5px; color: var(--sub); }
    footer.ps365-footer img { width: 18px; height: 18px; }
    footer.ps365-footer a { color: var(--text); font-weight: 600; text-decoration: none; }
    footer.ps365-footer a.clidsys { color: var(--brand-solid); }
    footer.ps365-footer a:hover { text-decoration: underline; text-underline-offset: 2px; }
    footer.ps365-footer .spacer { margin-left: auto; }
'@

    $logoImg = if ($logoDataUri) { "<img src='$logoDataUri' alt=''>" } else { '' }

    $headerBrandHtml = @"
<div class="ps365-brand">
    $logoImg
    <span class="ps365-brand-name">PS365</span>
    <span class="ps365-brand-by">by <a href="$clidsysUrl" target="_blank" rel="noopener">Clidsys</a></span>
  </div>
"@

    $encodedTitle = [System.Net.WebUtility]::HtmlEncode($Title)
    $subtitleHtml = if ($Subtitle) { "<p class=""page-sub"">$([System.Net.WebUtility]::HtmlEncode($Subtitle))</p>" } else { '' }
    $actionsBlock = if ($ActionsHtml) { "<div class=""page-actions"">$ActionsHtml</div>" } else { '' }
    $pageHeadHtml = "<div class=""page-head""><div><h1>$encodedTitle</h1>$subtitleHtml</div>$actionsBlock</div>"

    $footerHtml = @"
<footer class="ps365-footer">
  $logoImg
  <span>Generated by <a href="$ps365Url" target="_blank" rel="noopener">PS365</a>$versionText on $generatedOn</span>
  <span class="spacer">Built and maintained by <a class="clidsys" href="$clidsysUrl" target="_blank" rel="noopener">Clidsys</a></span>
</footer>
"@

    # Browser tab icon: the PS365 logo, also avoids the favicon 404 in the console
    $headHtml = if ($logoDataUri) { "<link rel=""icon"" type=""image/svg+xml"" href=""$logoDataUri"">" } else { '' }

    # Copy button markup, reused by every section that offers a copy. The button copies the closest element
    # carrying the data-copy-root attribute
    $copyButtonHtml = '<button class="btn btn-copy" type="button" onclick="event.preventDefault(); event.stopPropagation(); ps365CopyBlock(this)" title="Copy this block (pastes as a table in Word)"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><rect x="9" y="9" width="11" height="11" rx="2"></rect><path d="M5 15V5a2 2 0 0 1 2-2h10"></path></svg><span>Copy</span></button>'

    # Copy logic, shared by the pages. The copied block is a cleaned clone: screen-only parts removed (copy
    # buttons, search boxes, the stacked bar, anything marked no-copy), every collapsible section expanded,
    # and inline borders on the tables, so a paste into Word or Outlook keeps readable tables (a paste drops
    # the page stylesheet). Rich HTML through the Clipboard API, with a select-and-copy fallback when the
    # browser refuses it.
    $scriptHtml = @'
<script>
  function ps365CopyBlock(button) {
    var root = button.closest('[data-copy-root]');
    if (!root) { return; }
    // Wrapped so the root itself is processed too: a copied site row is a closed <details>, whose
    // content would otherwise stay hidden and copy as nothing
    var clone = document.createElement('div');
    clone.appendChild(root.cloneNode(true));
    clone.firstChild.removeAttribute('style');
    clone.querySelectorAll('.btn-copy, .no-copy, input, script, details.affected-list > summary').forEach(function (n) { n.remove(); });
    clone.querySelectorAll('[style*="display: none"], [style*="display:none"]').forEach(function (n) { n.remove(); });
    clone.querySelectorAll('summary, details').forEach(function (n) {
      var div = document.createElement('div');
      div.innerHTML = n.innerHTML;
      n.replaceWith(div);
    });
    clone.querySelectorAll('table').forEach(function (t) {
      t.setAttribute('border', '1');
      t.setAttribute('cellpadding', '4');
      t.style.borderCollapse = 'collapse';
      t.style.fontFamily = 'Segoe UI, sans-serif';
      t.style.fontSize = '10pt';
    });
    clone.querySelectorAll('th').forEach(function (c) { c.style.background = '#F4F5F7'; c.style.textAlign = 'left'; });

    var holder = document.createElement('div');
    holder.style.position = 'fixed';
    holder.style.left = '-10000px';
    holder.style.top = '0';
    holder.appendChild(clone);
    document.body.appendChild(holder);
    var html = holder.innerHTML;
    var text = holder.innerText;

    function done(ok) {
      holder.remove();
      var label = button.querySelector('span');
      if (label) { label.textContent = ok ? 'Copied' : 'Copy failed'; }
      button.classList.toggle('done', ok);
      setTimeout(function () { if (label) { label.textContent = 'Copy'; } button.classList.remove('done'); }, 1600);
    }

    function fallback() {
      var range = document.createRange();
      range.selectNodeContents(holder);
      var selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
      var ok = false;
      try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
      selection.removeAllRanges();
      done(ok);
    }

    if (navigator.clipboard && window.ClipboardItem) {
      var item = new ClipboardItem({
        'text/html': new Blob([html], { type: 'text/html' }),
        'text/plain': new Blob([text], { type: 'text/plain' })
      });
      navigator.clipboard.write([item]).then(function () { done(true); }, fallback);
    }
    else {
      fallback();
    }
  }
</script>
'@

    return [PSCustomObject]@{
        Css             = $css
        HeadHtml        = $headHtml
        HeaderBrandHtml = $headerBrandHtml
        PageHeadHtml    = $pageHeadHtml
        FooterHtml      = $footerHtml
        CopyButtonHtml  = $copyButtonHtml
        ScriptHtml      = $scriptHtml
    }
}
