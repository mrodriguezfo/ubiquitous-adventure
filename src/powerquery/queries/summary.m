let
    Source = Excel.Workbook(File.Contents("S:\InfoCore\Aplicaciones\Modelos Información\Retornos\utilidades\utilidades - Octubre 25 - Test.xlsx"), null, true),
    summary_Sheet = Source{[Item="summary",Kind="Sheet"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(summary_Sheet,{{"Column1", type text}, {"Column2", type text}, {"Column3", type any}, {"Column4", type any}, {"Column5", type any}, {"Column6", type any}, {"Column7", type any}, {"Column8", type any}, {"Column9", type any}, {"Column10", type any}, {"Column11", type any}, {"Column12", type any}, {"Column13", type any}, {"Column14", type any}, {"Column15", type any}, {"Column16", type any}, {"Column17", type any}, {"Column18", type any}, {"Column19", type any}, {"Column20", type any}, {"Column21", type text}}),
    #"Removed Bottom Rows" = Table.RemoveLastN(#"Changed Type",29),
    #"Removed Columns" = Table.RemoveColumns(#"Removed Bottom Rows",{"Column1", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41"})
in
    #"Removed Columns"
