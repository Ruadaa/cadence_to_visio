; File: export_cdl_current.il
; Usage:
;   load("/path/to/cadence_to_visio/export_cdl_current.il")
;   export_cdl_current("/path/to/output/SA_PreCh.cdl")

procedure(export_cdl_current(outPath)
  let((cv libName cellName viewName tmpDir tmpTag cdsLib cmd tmpNetlist outFile)

    unless(outPath
      error("export_cdl_current: outPath is required.\n")
    )
    outFile = outPath

    cv = geGetEditCellView()
    unless(cv
      error("No current edit cellview. Open a schematic window first.\n")
    )

    libName  = cv~>libName
    cellName = cv~>cellName
    viewName = cv~>viewName

    unless(viewName == "schematic"
      error("Current cellview is %s/%s/%s, not schematic.\n"
            libName cellName viewName)
    )

    schCheck(cv)
    dbSave(cv)

    ; Make sure output directory exists (shell dirname, avoids SKILL dirname)
    sh(sprintf(nil "mkdir -p \"`dirname '%s'`\"" outFile))

    ; Always use a fresh run dir to avoid the simInitEnvWithArgs dialog.
    ; tmpTag uses current time (seconds). If two calls happen within 1s, the
    ; rm -rf ensures it's still clean.
    tmpTag = buildString(parseString(buildString(parseString(getCurrentTime()) "_") ":") "_")
    tmpDir = strcat("/tmp/zz_cdlout_" libName "_" cellName "_" tmpTag)

    ; Ensure it does NOT exist. Do NOT mkdir it (simInitEnvWithArgs will).
    sh(sprintf(nil "rm -rf '%s'" tmpDir))

    ; Generate si.env for auCdl netlisting (may create tmpDir)
    simInitEnvWithArgs(tmpDir libName cellName viewName "auCdl" nil)

    cdsLib = simplifyFilename("./cds.lib")

    if(cdsLib && cdsLib != "" then
      cmd = sprintf(nil
        "cd '%s' ; si -batch -cdslib '%s' -command netlist > /dev/null 2>&1"
        tmpDir cdsLib)
      else
      cmd = sprintf(nil
        "cd '%s' ; si -batch -command netlist > /dev/null 2>&1"
        tmpDir)
    )

    sh(cmd)

    tmpNetlist = strcat(tmpDir "/netlist")
    unless(isFile(tmpNetlist)
      sh(sprintf(nil "rm -rf '%s'" tmpDir))
      error("CDL netlist not generated. (si failed)\n")
    )

    sh(sprintf(nil "cp '%s' '%s'" tmpNetlist outFile))

    ; Cleanup intermediates
    sh(sprintf(nil "rm -rf '%s'" tmpDir))

    printf("CDL netlist exported to: %s\n" outFile)
    outFile
  )
)
