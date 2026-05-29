; File: sch2visio.il
; Purpose: One-shot export for cadence_to_visio
; Output files (under outDir):
;   netlist.cdl
;   inst_info.txt
;   wires.tsv
;
; Usage:
;   load("/path/to/cadence_to_visio/sch2visio.il")
;   sch2visio("/path/to/output/out_sa_prech")

procedure(sch2visio(outDir)
  let((cdlIl instIl wireIl cdlOut instOut wireOut)

    unless(outDir
      error("sch2visio: outDir is required.\n")
    )

    ; Ensure outDir exists
    sh(sprintf(nil "mkdir -p '%s'" outDir))

    ; Paths to existing exporter scripts
    cdlIl  = "/path/to/cadence_to_visio/export_cdl_current.il"
    instIl = "/path/to/cadence_to_visio/export_inst_xy_orient.il"
    wireIl = "/path/to/cadence_to_visio/export_schematic_wires.il"

    ; Load exporters
    load(cdlIl)
    load(instIl)
    load(wireIl)

    ; Sanity checks (fail early with clear message)
    unless(fboundp('export_cdl_current)
      error("sch2visio: export_cdl_current not found after load(%s)\n" cdlIl)
    )
    unless(fboundp('c2vExportInstXYOrient)
      error("sch2visio: c2vExportInstXYOrient not found after load(%s)\n" instIl)
    )
    unless(fboundp('c2vExportWireLinesV4)
      error("sch2visio: c2vExportWireLinesV4 not found after load(%s)\n" wireIl)
    )

    ; Output paths
    cdlOut  = strcat(outDir "/netlist.cdl")
    instOut = strcat(outDir "/inst_info.txt")
    wireOut = strcat(outDir "/wires.tsv")

    ; Run exports (current schematic must be the active edit window)
    export_cdl_current(cdlOut)
    c2vExportInstXYOrient(instOut)
    c2vExportWireLinesV4(wireOut)

    printf("sch2visio exported:\n  %s\n  %s\n  %s\n" cdlOut instOut wireOut)
    list(cdlOut instOut wireOut)
  )
)
