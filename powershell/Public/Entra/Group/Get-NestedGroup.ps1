<#
    .SYNOPSIS
    Lists all nested group dependencies (group-in-group) in Microsoft Entra ID.

    .DESCRIPTION
    The Get-NestedGroup function fetches all groups via the Microsoft Graph API,
    scans the members of each group to detect group-type members (nested groups),
    and outputs the dependencies. It supports optional export to Excel or CSV for
    further analysis.

    .PARAMETER ExportToExcel
    When specified, exports the results to an Excel file in the user's profile directory.
    Requires the ImportExcel module.

    .PARAMETER ExportToHtml
    When specified, exports the results to an interactive HTML graph file in the user's profile directory.
    The graph is rendered in the browser using D3.js (loaded from CDN).
    Nodes are color-coded by group type. Clicking a node zooms in and reveals its parent and child groups.
    The graph is frozen after stabilisation: individual nodes can be dragged without affecting the others.

    .PARAMETER NoPermissionCheck
    (Optional) Skip the Microsoft Graph scope verification performed against the current Get-MgContext token.

    .EXAMPLE
    Get-NestedGroup

    Retrieves all nested group dependencies and outputs them to the console.

    .EXAMPLE
    Get-NestedGroup -ExportToExcel

    Retrieves all nested group dependencies and exports the results to an Excel file.

    .EXAMPLE
    Get-NestedGroup -ExportToHtml

    Retrieves all nested group dependencies and exports an interactive HTML graph to the user's profile directory.

    .OUTPUTS
    System.Collections.Generic.List[Object]

    .NOTES
    OUTPUT PROPERTIES
    Returns a collection of custom objects with the following properties:
    - MemberGroup: Display name of the nested (child) group
    - MemberGroupId: Unique identifier of the nested (child) group
    - MemberGroupType: Type of the nested group (Microsoft 365, Dynamic, Mail-enabled Security, Security, Distribution, Other)
    - ParentGroup: Display name of the parent group containing the nested group
    - ParentGroupId: Unique identifier of the parent group
    - ParentGroupType: Type of the parent group

    Requires Microsoft.Graph module: Connect-MgGraph -Scopes 'Group.Read.All'

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-NestedGroup
#>

function Get-NestedGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToHtml,

        [Parameter(Mandatory = $false)]
        [switch]$NoPermissionCheck
    )

    if (-not $NoPermissionCheck.IsPresent) {
        $requiredScopes = @('Group.Read.All')
        if (-not (Test-MgGraphPermission -RequiredScopes $requiredScopes -CallerName $MyInvocation.MyCommand.Name)) {
            return
        }
    }

    Write-Verbose 'Fetching all groups...'
    Write-Host -ForegroundColor Cyan 'Fetching all groups...'

    # Fetch all groups with pagination
    [System.Collections.Generic.List[Object]]$allGroups = @()
    $uri = "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,groupTypes,securityEnabled,mailEnabled&`$top=999&`$count=true"
    $headers = @{ ConsistencyLevel = 'eventual' }

    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers
        foreach ($group in $response.value) {
            $allGroups.Add($group)
        }
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    # Build a lookup table for group names
    $groupLookup = @{}
    foreach ($group in $allGroups) {
        $groupLookup[$group.id] = $group
    }

    $totalCount = $allGroups.Count
    Write-Verbose "Found $totalCount groups. Scanning for nested group memberships..."
    Write-Host -ForegroundColor Cyan "Found $totalCount groups. Scanning for nested group memberships..."

    # Scan members of each group and find group-type members
    [System.Collections.Generic.List[Object]]$dependencies = @()
    $processed = 0

    foreach ($group in $allGroups) {
        $processed++
        if ($processed % 50 -eq 0) {
            Write-Progress -Activity 'Scanning group members' -Status "$processed / $totalCount" -PercentComplete (($processed / $totalCount) * 100)
        }

        try {
            $membersUri = "https://graph.microsoft.com/v1.0/groups/$($group.id)/members?`$select=id,displayName&`$top=999"
            $memberResponse = Invoke-MgGraphRequest -Method GET -Uri $membersUri

            do {
                foreach ($member in $memberResponse.value) {
                    # Check if the member is a group (has @odata.type = #microsoft.graph.group)
                    if ($member.'@odata.type' -eq '#microsoft.graph.group') {
                        $parentType = if ($group.groupTypes -contains 'Unified') { 'Microsoft 365' }
                                      elseif ($group.groupTypes -contains 'DynamicMembership') { 'Dynamic' }
                                      elseif ($group.securityEnabled -and $group.mailEnabled) { 'Mail-enabled Security' }
                                      elseif ($group.securityEnabled) { 'Security' }
                                      elseif ($group.mailEnabled) { 'Distribution' }
                                      else { 'Other' }

                        $childGroup = $groupLookup[$member.id]
                        $childType = if ($childGroup) {
                            if ($childGroup.groupTypes -contains 'Unified') { 'Microsoft 365' }
                            elseif ($childGroup.groupTypes -contains 'DynamicMembership') { 'Dynamic' }
                            elseif ($childGroup.securityEnabled -and $childGroup.mailEnabled) { 'Mail-enabled Security' }
                            elseif ($childGroup.securityEnabled) { 'Security' }
                            elseif ($childGroup.mailEnabled) { 'Distribution' }
                            else { 'Other' }
                        }
                        else { 'Unknown' }

                        $object = [PSCustomObject][ordered]@{
                            MemberGroup     = $member.displayName
                            MemberGroupId   = $member.id
                            MemberGroupType = $childType
                            ParentGroup     = $group.displayName
                            ParentGroupId   = $group.id
                            ParentGroupType = $parentType
                        }

                        $dependencies.Add($object)
                    }
                }

                $nextLink = $memberResponse.'@odata.nextLink'
                if ($nextLink) {
                    $memberResponse = Invoke-MgGraphRequest -Method GET -Uri $nextLink
                }
            } while ($nextLink)
        }
        catch {
            Write-Warning "Failed to scan members for group '$($group.displayName)': $_"
        }
    }

    Write-Progress -Activity 'Scanning group members' -Completed

    # Compute summary stats
    $uniqueGroupIds = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($dependency in $dependencies) {
        [void]$uniqueGroupIds.Add($dependency.MemberGroupId)
        [void]$uniqueGroupIds.Add($dependency.ParentGroupId)
    }
    $uniqueGroups = $uniqueGroupIds.Count

    Write-Host -ForegroundColor Yellow "`nNested Group Dependencies: $($dependencies.Count) relationships across $uniqueGroups groups`n"

    if ($dependencies.Count -eq 0) {
        Write-Host -ForegroundColor Green 'No nested group dependencies found.'
        return
    }

    if ($ExportToExcel.IsPresent) {
        Write-Verbose 'Preparing Excel export...'
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:USERPROFILE)\$now-NestedGroups.xlsx"
        Write-Verbose "Excel file path: $excelFilePath"
        Write-Host -ForegroundColor Cyan "Exporting nested groups to Excel file: $excelFilePath"
        $dependencies | Sort-Object ParentGroup, MemberGroup | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'NestedGroups'
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    elseif ($ExportToHtml.IsPresent) {
        Write-Verbose 'Preparing HTML graph export...'
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $htmlFilePath = "$($env:USERPROFILE)\$now-NestedGroups.html"
        $generatedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $relCount = $dependencies.Count

        # Build unique node collection from all group IDs referenced in dependencies
        $nodeMap = @{}
        foreach ($dep in $dependencies) {
            if (-not $nodeMap.ContainsKey($dep.MemberGroupId)) {
                $nodeMap[$dep.MemberGroupId] = [PSCustomObject]@{
                    id    = $dep.MemberGroupId
                    label = $dep.MemberGroup
                    group = $dep.MemberGroupType
                    title = "$($dep.MemberGroup) [$($dep.MemberGroupType)]"
                }
            }
            if (-not $nodeMap.ContainsKey($dep.ParentGroupId)) {
                $nodeMap[$dep.ParentGroupId] = [PSCustomObject]@{
                    id    = $dep.ParentGroupId
                    label = $dep.ParentGroup
                    group = $dep.ParentGroupType
                    title = "$($dep.ParentGroup) [$($dep.ParentGroupType)]"
                }
            }
        }

        [array]$nodeList = $nodeMap.Values
        [array]$edgeList = $dependencies | ForEach-Object {
            [PSCustomObject]@{ from = $_.MemberGroupId; to = $_.ParentGroupId }
        }

        $nodeCount     = $nodeList.Count
        $nodesJson     = ConvertTo-Json -InputObject $nodeList -Compress -Depth 3
        $edgesJson     = ConvertTo-Json -InputObject $edgeList -Compress -Depth 3
        $statsText     = "$nodeCount groups  - $relCount nested relationships"

        # Single-quoted here-string: no PS variable expansion inside HTML/JS.
        # Placeholders are replaced via .NET String.Replace() (literal, not regex).
        $htmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Nested Groups - PS365</title>
  <script>PS365_D3_LIB</script>
  <style>
    :root {
      --bg: #f4f6f9; --canvas: #ffffff; --surface: #ffffff; --surface2: #f0f2f5;
      --border: #dfe3e8; --text: #1e2a3a; --muted: #6b7a8d; --sub: #8995a5;
      --accent: #2563eb; --accent-light: #dbeafe; --accent2: #7c3aed;
      --radius: 8px; --shadow: 0 1px 3px rgba(0,0,0,.06), 0 1px 2px rgba(0,0,0,.04);
    }
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: "Segoe UI", system-ui, -apple-system, sans-serif; background: var(--bg); color: var(--text); display: flex; flex-direction: column; height: 100vh; overflow: hidden; }

    /* ── HEADER ── */
    header { background: var(--surface); border-bottom: 1px solid var(--border); padding: 0 24px; height: 56px; display: flex; align-items: center; gap: 18px; flex-shrink: 0; box-shadow: var(--shadow); z-index: 10; }
    .brand { display: flex; align-items: center; gap: 11px; flex-shrink: 0; }
    .brand-icon { width: 32px; height: 32px; background: linear-gradient(135deg, var(--accent), var(--accent2)); border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 16px; color: #fff; }
    .brand-title { font-size: 15px; font-weight: 650; color: var(--text); }
    .brand-sub { font-size: 11px; color: var(--sub); }
    .sep { width: 1px; height: 28px; background: var(--border); flex-shrink: 0; }
    .chip { background: var(--accent-light); border: 1px solid #bfdbfe; border-radius: 20px; padding: 3px 12px; font-size: 11.5px; color: var(--accent); font-weight: 600; white-space: nowrap; flex-shrink: 0; }
    .chip.muted { background: var(--surface2); border-color: var(--border); color: var(--muted); font-weight: 400; }
    .search-wrap { position: relative; flex: 0 0 210px; }
    .search-icon { position: absolute; left: 10px; top: 50%; transform: translateY(-50%); color: var(--muted); font-size: 13px; pointer-events: none; }
    #search { width: 100%; background: var(--surface2); border: 1px solid var(--border); border-radius: 20px; color: var(--text); padding: 6px 12px 6px 30px; font-size: 12.5px; outline: none; transition: border-color .2s, box-shadow .2s; }
    #search:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(37,99,235,.12); }
    #search::placeholder { color: var(--sub); }
    .toolbar { display: flex; gap: 6px; }
    .btn { background: var(--surface); border: 1px solid var(--border); color: var(--text); padding: 5px 14px; border-radius: var(--radius); font-size: 12px; font-family: inherit; cursor: pointer; transition: background .15s, border-color .15s, box-shadow .15s; white-space: nowrap; }
    .btn:hover { background: var(--surface2); border-color: #c4cad2; box-shadow: var(--shadow); }

    /* ── FILTER BAR ── */
    .filter-bar { background: var(--surface); border-bottom: 1px solid var(--border); padding: 6px 24px; display: flex; align-items: center; gap: 8px; flex-shrink: 0; flex-wrap: wrap; }
    .filter-lbl { font-size: 10.5px; text-transform: uppercase; letter-spacing: .6px; font-weight: 600; color: var(--sub); margin-right: 4px; }
    .filter-item { display: flex; align-items: center; gap: 5px; font-size: 12px; color: var(--text); cursor: pointer; padding: 3px 10px; border-radius: 6px; border: 1px solid var(--border); background: var(--surface); transition: background .12s, border-color .12s; user-select: none; }
    .filter-item:hover { background: var(--surface2); }
    .filter-item.off { opacity: .45; }
    .filter-item input { cursor: pointer; accent-color: var(--accent); margin: 0; width: 13px; height: 13px; }
    .filter-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }

    /* ── MAIN ── */
    #wrap { display: flex; flex: 1; overflow: hidden; }
    #network { flex: 1; display: block; background: var(--canvas); border-radius: 0; }

    /* ── PANEL ── */
    #panel { width: 310px; background: var(--surface); border-left: 1px solid var(--border); display: none; flex-direction: column; overflow: hidden; box-shadow: -2px 0 8px rgba(0,0,0,.04); }
    .panel-header { padding: 18px 18px 14px; border-bottom: 1px solid var(--border); display: flex; align-items: flex-start; gap: 10px; }
    .panel-type-dot { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; margin-top: 5px; }
    .panel-name { font-size: 14px; font-weight: 650; line-height: 1.45; word-break: break-word; flex: 1; }
    .panel-close { flex-shrink: 0; width: 24px; height: 24px; border-radius: 6px; background: transparent; border: none; color: var(--muted); font-size: 17px; cursor: pointer; display: flex; align-items: center; justify-content: center; transition: background .15s, color .15s; }
    .panel-close:hover { background: var(--surface2); color: var(--text); }
    .panel-body { padding: 16px 18px; overflow-y: auto; flex: 1; }
    .meta-row { display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 10px; padding-bottom: 10px; border-bottom: 1px solid var(--border); }
    .meta-lbl { font-size: 11px; color: var(--sub); }
    .meta-val { font-size: 12.5px; font-weight: 600; }
    .id-row { background: var(--surface2); border: 1px solid var(--border); border-radius: var(--radius); padding: 8px 11px; margin-bottom: 16px; }
    .id-lbl { font-size: 10px; color: var(--sub); margin-bottom: 3px; letter-spacing: .5px; text-transform: uppercase; font-weight: 600; }
    .id-val { font-size: 10.5px; color: var(--muted); font-family: "Cascadia Code","Consolas",monospace; word-break: break-all; line-height: 1.5; }
    .sec { margin-bottom: 18px; }
    .sec-title { font-size: 10.5px; color: var(--sub); text-transform: uppercase; letter-spacing: .6px; font-weight: 700; margin-bottom: 8px; display: flex; align-items: center; gap: 7px; }
    .sec-title::after { content: ""; flex: 1; height: 1px; background: var(--border); }
    .badge { display: inline-flex; align-items: center; gap: 5px; padding: 5px 11px; border-radius: 6px; font-size: 11.5px; margin: 3px 3px 0 0; cursor: pointer; color: #fff; font-weight: 600; transition: opacity .15s, transform .1s; line-height: 1.3; box-shadow: 0 1px 2px rgba(0,0,0,.1); }
    .badge:hover { opacity: .88; transform: translateY(-1px); }
    .badge-dot { width: 6px; height: 6px; border-radius: 50%; background: rgba(255,255,255,.55); flex-shrink: 0; }
  </style>
</head>
<body>

<header>
  <div class="brand">
      <div class="brand-title">Nested Group Dependencies</div>
      <div class="brand-sub">From <a href="https://ps365.clidsys.com" target="_blank" rel="noopener" style="color:inherit;text-decoration:underline;text-underline-offset:2px">PS365</a></div>
    </div>
  </div>
  <div class="sep"></div>
  <span class="chip">PS365_STATS_TEXT</span>
  <span class="chip muted">PS365_GENERATED_DATE</span>
  <div class="search-wrap">
    <input id="search" type="text" placeholder="Search groups&hellip;" autocomplete="off" />
  </div>
  <div class="toolbar">
    <button class="btn" onclick="fitGraph()">&#8982; Fit</button>
    <button class="btn" onclick="resetSearch()">&#10006; Reset</button>
    <button class="btn" onclick="exportPNG()">&#x2197; PNG</button>
  </div>
</header>

<div class="filter-bar" id="filter-bar">
  <span class="filter-lbl">Filter</span>
</div>

<div id="wrap">
  <svg id="network"></svg>
  <div id="panel">
    <div class="panel-header">
      <div class="panel-type-dot" id="p-dot"></div>
      <div class="panel-name" id="p-name"></div>
      <button class="panel-close" onclick="document.getElementById('panel').style.display='none'" title="Close">&#xD7;</button>
    </div>
    <div class="panel-body">
      <div class="meta-row">
        <span class="meta-lbl">Type</span>
        <span class="meta-val" id="p-type"></span>
      </div>
      <div class="id-row">
        <div class="id-lbl">Object ID</div>
        <div class="id-val" id="p-id"></div>
      </div>
      <div class="sec" id="sec-parents" style="display:none">
        <div class="sec-title">Member of</div>
        <div id="p-parents"></div>
      </div>
      <div class="sec" id="sec-children" style="display:none">
        <div class="sec-title">Contains</div>
        <div id="p-children"></div>
      </div>
    </div>
  </div>
</div>

<script>
  var nodesData = PS365_NODES_JSON;
  var edgesData = PS365_EDGES_JSON;

  var colorMap = {
    "Microsoft 365":         { bg: "#2563eb", border: "#3b82f6", node: "#1d4ed8", text: "#fff" },
    "Security":              { bg: "#dc2626", border: "#ef4444", node: "#b91c1c", text: "#fff" },
    "Dynamic":               { bg: "#16a34a", border: "#22c55e", node: "#15803d", text: "#fff" },
    "Distribution":          { bg: "#7c3aed", border: "#8b5cf6", node: "#6d28d9", text: "#fff" },
    "Mail-enabled Security": { bg: "#d97706", border: "#f59e0b", node: "#b45309", text: "#fff" },
    "Other":                 { bg: "#64748b", border: "#94a3b8", node: "#475569", text: "#fff" },
    "Unknown":               { bg: "#94a3b8", border: "#cbd5e1", node: "#64748b", text: "#fff" }
  };
  function cData(g)      { return colorMap[g] || colorMap["Other"]; }
  function colorBg(g)    { return cData(g).node; }
  function colorBorder(g){ return cData(g).border; }

  // ── BUILD FILTER CHECKBOXES ──
  var usedTypes = {};
  nodesData.forEach(function(n) { usedTypes[n.group] = true; });
  var filterBar = document.getElementById("filter-bar");
  Object.keys(colorMap).filter(function(t) { return t !== "Unknown" && usedTypes[t]; }).forEach(function(t) {
    var lbl = document.createElement("label");
    lbl.className = "filter-item";
    lbl.innerHTML = '<input type="checkbox" checked data-type="' + t + '">' +
                    '<span class="filter-dot" style="background:' + cData(t).border + '"></span>' + t;
    lbl.querySelector("input").addEventListener("change", applyFilter);
    filterBar.appendChild(lbl);
  });

  // ── SVG SETUP ──
  var wrap = document.getElementById("wrap");
  var svgEl = document.getElementById("network");
  var W = wrap.clientWidth, H = wrap.clientHeight;
  var svg = d3.select(svgEl).attr("width", W).attr("height", H);

  var defs = svg.append("defs");
  // Drop-shadow filter (light mode)
  var filt = defs.append("filter").attr("id", "node-shadow").attr("x", "-20%").attr("y", "-20%").attr("width", "140%").attr("height", "140%");
  filt.append("feDropShadow").attr("dx", 0).attr("dy", 1).attr("stdDeviation", 2).attr("flood-opacity", 0.10);

  // Arrow markers
  Object.keys(colorMap).forEach(function(g) {
    defs.append("marker")
      .attr("id", "arr-" + g.replace(/ /g, "_"))
      .attr("viewBox", "0 -5 10 10").attr("refX", 10).attr("refY", 0)
      .attr("markerWidth", 6).attr("markerHeight", 6).attr("orient", "auto")
      .attr("overflow", "visible")
      .append("path").attr("d", "M0,-5L10,0L0,5").attr("fill", colorBorder(g));
  });

  var zoomG = svg.append("g");
  var zoomBehavior = d3.zoom().scaleExtent([0.05, 8]).on("zoom", function(event) { zoomG.attr("transform", event.transform); });
  svg.call(zoomBehavior);
  svg.on("click", function(event) { if (event.target === svgEl) document.getElementById("panel").style.display = "none"; });

  // ── SIMULATION ──
  var byId = {};
  var simNodes = nodesData.map(function(n) { var s = { id: n.id, label: n.label, group: n.group, visible: true }; byId[n.id] = s; return s; });
  var simLinks = edgesData.map(function(e) { return { source: e.from, target: e.to, srcGroup: byId[e.from] ? byId[e.from].group : "Other" }; });

  var sim = d3.forceSimulation(simNodes)
    .force("link",      d3.forceLink(simLinks).id(function(d) { return d.id; }).distance(160))
    .force("charge",    d3.forceManyBody().strength(-700))
    .force("center",    d3.forceCenter(W / 2, H / 2))
    .force("collision", d3.forceCollide(70));

  // ── EDGES ──
  var link = zoomG.append("g").selectAll("line").data(simLinks).enter().append("line")
    .attr("stroke",         function(d) { return colorBorder(d.srcGroup); })
    .attr("stroke-width",   1.4)
    .attr("stroke-opacity", 0.5)
    .attr("marker-end",     function(d) { return "url(#arr-" + d.srcGroup.replace(/ /g, "_") + ")"; });

  // ── NODES ──
  var NODE_W = 168, NODE_H = 30;
  var node = zoomG.append("g").selectAll("g").data(simNodes).enter().append("g")
    .style("cursor", "pointer")
    .call(d3.drag()
      .on("start", function(event, d) {
        if (!event.active) sim.alphaTarget(0.2).restart();
        d.fx = d.x; d.fy = d.y;
      })
      .on("drag",  function(event, d) { d.fx = event.x; d.fy = event.y; })
      .on("end",   function(event, d) {
        if (!event.active) sim.alphaTarget(0);
        d.fx = event.x; d.fy = event.y;
      })
    )
    .on("click", function(event, d) { event.stopPropagation(); selectNode(d.id); });

  node.append("rect")
    .attr("width", NODE_W).attr("height", NODE_H)
    .attr("x", -NODE_W / 2).attr("y", -NODE_H / 2)
    .attr("rx", 6)
    .attr("fill",         function(d) { return colorBg(d.group); })
    .attr("stroke",       function(d) { return colorBorder(d.group); })
    .attr("stroke-width", 1.2)
    .attr("filter",       "url(#node-shadow)");

  node.append("text")
    .text(function(d) { return d.label.length > 22 ? d.label.substring(0, 21) + "\u2026" : d.label; })
    .attr("text-anchor", "middle").attr("dy", "0.35em")
    .attr("fill", "#fff")
    .style("font-size", "11.5px").style("font-weight", "500")
    .style("font-family", "Segoe UI, system-ui, sans-serif")
    .style("pointer-events", "none");

  // Clip line endpoints to rectangle edges so arrows land exactly at node borders
  var HW = NODE_W / 2 + 3, HH = NODE_H / 2 + 3; // +3px gap
  function clipPt(ox, oy, tx, ty) {
    var dx = tx - ox, dy = ty - oy;
    if (Math.abs(dx) < 0.01 && Math.abs(dy) < 0.01) return { x: tx, y: ty };
    var rx = Math.abs(dx) > 0 ? HW / Math.abs(dx) : Infinity;
    var ry = Math.abs(dy) > 0 ? HH / Math.abs(dy) : Infinity;
    var r = Math.min(rx, ry, 1);
    return { x: tx - dx * r, y: ty - dy * r };
  }

  function ticked() {
    link.each(function(d) {
      var s = clipPt(d.target.x, d.target.y, d.source.x, d.source.y);
      var e = clipPt(d.source.x, d.source.y, d.target.x, d.target.y);
      d3.select(this).attr("x1", s.x).attr("y1", s.y).attr("x2", e.x).attr("y2", e.y);
    });
    node.attr("transform", function(d) { return "translate(" + d.x + "," + d.y + ")"; });
  }

  sim.on("tick", ticked).on("end", function() { simNodes.forEach(function(d) { d.fx = d.x; d.fy = d.y; }); });

  // ── FILTER ──
  function applyFilter() {
    var hidden = {};
    document.querySelectorAll("#filter-bar input[type=checkbox]").forEach(function(cb) {
      var item = cb.closest(".filter-item");
      if (!cb.checked) { hidden[cb.getAttribute("data-type")] = true; item.classList.add("off"); }
      else { item.classList.remove("off"); }
    });
    node.style("display", function(d) { d.visible = !hidden[d.group]; return d.visible ? null : "none"; });
    link.style("display", function(d) {
      var src = typeof d.source === "object" ? d.source : byId[d.source];
      var tgt = typeof d.target === "object" ? d.target : byId[d.target];
      return (src && src.visible && tgt && tgt.visible) ? null : "none";
    });
  }

  // ── FIT ──
  function fitGraph() {
    var visible = simNodes.filter(function(d) { return d.visible !== false; });
    if (!visible.length) return;
    var xs = visible.map(function(d) { return d.x; }), ys = visible.map(function(d) { return d.y; });
    var minX = d3.min(xs) - NODE_W, maxX = d3.max(xs) + NODE_W;
    var minY = d3.min(ys) - NODE_H * 2, maxY = d3.max(ys) + NODE_H * 2;
    var scale = Math.min(0.92 * W / (maxX - minX), 0.92 * H / (maxY - minY), 3);
    svg.transition().duration(500).call(zoomBehavior.transform, d3.zoomIdentity.translate((W - scale * (minX + maxX)) / 2, (H - scale * (minY + maxY)) / 2).scale(scale));
  }

  // ── SELECT + PANEL ──
  function selectNode(id) {
    var n = byId[id];
    if (n && n.x !== undefined) {
      svg.transition().duration(400).call(zoomBehavior.transform, d3.zoomIdentity.translate(W / 2 - 1.4 * n.x, H / 2 - 1.4 * n.y).scale(1.4));
    }
    showPanel(id);
  }

  function showPanel(id) {
    var nd = nodesData.find(function(n) { return n.id === id; });
    if (!nd) return;
    document.getElementById("p-name").textContent = nd.label;
    document.getElementById("p-type").textContent = nd.group;
    document.getElementById("p-id").textContent   = nd.id;
    document.getElementById("p-dot").style.background = colorBorder(nd.group);
    var parents  = edgesData.filter(function(e) { return e.from === id; });
    var children = edgesData.filter(function(e) { return e.to   === id; });
    function makeBadge(otherId) {
      var on = nodesData.find(function(x) { return x.id === otherId; });
      var col = colorBorder(on ? on.group : "Other");
      var bg  = colorBg(on ? on.group : "Other");
      return "<span class='badge' style='background:" + bg + ";border:1px solid " + col + "'" +
             " onclick=\"selectNode('" + otherId + "')\">" +
             "<span class='badge-dot' style='background:" + col + "'></span>" +
             (on ? on.label : otherId) + "</span>";
    }
    var secP = document.getElementById("sec-parents"), secC = document.getElementById("sec-children");
    if (parents.length > 0) { document.getElementById("p-parents").innerHTML = parents.map(function(e) { return makeBadge(e.to); }).join(""); secP.style.display = "block"; }
    else { secP.style.display = "none"; }
    if (children.length > 0) { document.getElementById("p-children").innerHTML = children.map(function(e) { return makeBadge(e.from); }).join(""); secC.style.display = "block"; }
    else { secC.style.display = "none"; }
    document.getElementById("panel").style.display = "flex";
  }

  // ── SEARCH ──
  function resetSearch() {
    document.getElementById("search").value = "";
    node.select("rect").attr("fill", function(d) { return colorBg(d.group); }).attr("stroke", function(d) { return colorBorder(d.group); }).attr("stroke-width", 1.2);
    node.style("opacity", 1); link.attr("stroke-opacity", 0.5);
  }

  document.getElementById("search").addEventListener("input", function() {
    var val = this.value.trim().toLowerCase();
    if (!val) { resetSearch(); return; }
    var matchIds = {};
    nodesData.forEach(function(n) { if (n.label.toLowerCase().indexOf(val) !== -1) matchIds[n.id] = true; });
    node.style("opacity", function(d) { return matchIds[d.id] ? 1 : 0.08; });
    node.select("rect").attr("stroke", function(d) { return matchIds[d.id] ? "#1e293b" : colorBorder(d.group); }).attr("stroke-width", function(d) { return matchIds[d.id] ? 2.5 : 1; });
    link.attr("stroke-opacity", 0.04);
    var matched = simNodes.filter(function(d) { return matchIds[d.id]; });
    if (matched.length > 0) {
      var xs = matched.map(function(d) { return d.x; }), ys = matched.map(function(d) { return d.y; });
      var scale = Math.min(0.9 * W / (d3.max(xs) - d3.min(xs) + NODE_W * 4), 0.9 * H / (d3.max(ys) - d3.min(ys) + NODE_H * 8), 2);
      svg.transition().duration(500).call(zoomBehavior.transform, d3.zoomIdentity.translate((W - scale * (d3.min(xs) + d3.max(xs))) / 2, (H - scale * (d3.min(ys) + d3.max(ys))) / 2).scale(scale));
    }
  });

  // ── EXPORT PNG ──
  function exportPNG() {
    var clone = svgEl.cloneNode(true);
    clone.setAttribute("xmlns", "http://www.w3.org/2000/svg");
    // Apply current zoom transform to the clone
    var g = clone.querySelector("g");
    if (g) { g.setAttribute("transform", zoomG.attr("transform") || ""); }
    // Inline computed styles for proper rendering
    clone.querySelectorAll("text").forEach(function(t) {
      t.setAttribute("font-family", "Segoe UI, system-ui, sans-serif");
      t.setAttribute("font-size", "11.5px");
      t.setAttribute("font-weight", "500");
    });
    var xml = new XMLSerializer().serializeToString(clone);
    var blob = new Blob([xml], { type: "image/svg+xml;charset=utf-8" });
    var url = URL.createObjectURL(blob);
    var img = new Image();
    img.onload = function() {
      var scale = 2;
      var canvas = document.createElement("canvas");
      canvas.width = W * scale; canvas.height = H * scale;
      var ctx = canvas.getContext("2d");
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      URL.revokeObjectURL(url);
      var a = document.createElement("a");
      a.download = "nested-groups.png";
      a.href = canvas.toDataURL("image/png");
      a.click();
    };
    img.src = url;
  }
</script>
</body>
</html>
'@

        # Read D3.js library from Private/js folder
        $d3LibPath = Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\Private\js\d3.v7.min.js'
        $d3Content = Get-Content -Path $d3LibPath -Raw

        # Use .NET String.Replace (literal, no regex) to safely inject JSON data and metadata
        $html = $htmlTemplate
        $html = $html.Replace('PS365_D3_LIB',         $d3Content)
        $html = $html.Replace('PS365_NODES_JSON',     $nodesJson)
        $html = $html.Replace('PS365_EDGES_JSON',     $edgesJson)
        $html = $html.Replace('PS365_GENERATED_DATE', $generatedDate)
        $html = $html.Replace('PS365_STATS_TEXT',     $statsText)

        Write-Verbose "HTML graph file path: $htmlFilePath"
        $html | Out-File -FilePath $htmlFilePath -Encoding UTF8
        Write-Host -ForegroundColor Green "HTML graph exported to: $htmlFilePath"
        Invoke-Item -Path $htmlFilePath
    }
    else {
        Write-Verbose "Returning $($dependencies.Count) nested group dependencies"
        return $dependencies
    }
}
