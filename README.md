准备工作：
visio软件
python            pip install pywin32
netlist.txt       File -> export ->CDL
inst_info.txt     instances坐标信息 

#skill
procedure( exportInstXYOrient(cv outFile)
  let( (fp)
    fp = outfile(outFile "w")
    foreach(inst cv~>instances
      fprintf(fp "Name: %s  Cell: %s\n" inst~>name inst~>cellName)
      fprintf(fp "  XY: %L\n" inst~>xy)
      fprintf(fp "  Orient: %s\n" inst~>orient)
      fprintf(fp "  BBox: %L\n\n" inst~>bBox)
    )
    close(fp)
  )
)
exportInstXYOrient( geGetEditCellView() "/home/.../inst_info.txt" )

5.运行demo1.py
  
