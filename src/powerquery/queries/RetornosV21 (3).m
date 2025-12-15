let
    Source = Csv.Document(File.Contents("S:\InfoCore\Aplicaciones\Modelos Información\Retornos\RetornosV21.csv"),[Delimiter=",", Columns=17, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", type text}, {"Column2", type text}, {"Column3", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type text}, {"Column7", type text}, {"Column8", type text}, {"Column9", type text}, {"Column10", type text}, {"Column11", type text}, {"Column12", type text}, {"Column13", type text}, {"Column14", type text}, {"Column15", type text}, {"Column16", type text}, {"Column17", type text}}),
    #"Removed Top Rows" = Table.Skip(#"Changed Type",5),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Promoted Headers",{{"Account", type text}, {"Begin Date", type date}, {"End Date", type date}, {"Perf. Class", type text}, {"Settlement Date Cash Balance", type number}, {"Total Market Value", type number}, {"TWRR", type number}, {"TWRR M-T-D", type number}, {"TWRR Y-T-D", type number}, {"TWRR 3 month", type number}, {"TWRR 1 yr.", type number}, {"TWRR 3 yr. Ann.", type number}, {"TWRR Incept. Ann.", type number}, {"TWRR w/ Fees", type number}, {"TWRR w/Fees M-T-D", type number}, {"TWRR w/Fees Y-T-D", type number}, {"Total Earnings", type number}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type1", each Date.IsInPreviousNDays([Begin Date], 65)),
    #"Filtered Rows1" = Table.SelectRows(#"Filtered Rows", each ([Account] <> "INTACT+CALL" and [Account] <> "INTER-PASCOMPOS"))
in
    #"Filtered Rows1"
