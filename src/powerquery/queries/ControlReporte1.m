let
    Source = Csv.Document(File.Contents("S:\InfoCore\Aplicaciones\Modelos Información\Retornos\ControlReporte1.csv"),[Delimiter=",", Columns=10, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", type text}, {"Column2", type text}, {"Column3", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type text}, {"Column7", type text}, {"Column8", type text}, {"Column9", type text}, {"Column10", type text}}),
    #"Removed Top Rows" = Table.Skip(#"Changed Type",5),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Promoted Headers",{{"Description", type text}, {"Begin Total Market Value", type number}, {"End Total Market Value", type number}, {"Daily Total TWRR", type number}, {"Daily Total TWRR Tot. Contrib.", type number}, {"", type text}, {"_1", type text}, {"_2", type text}, {"_3", type text}, {"_4", type text}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{"", "_1", "_2", "_3", "_4"})
in
    #"Removed Columns"
