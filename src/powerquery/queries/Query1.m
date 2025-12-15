let
    Source = Sql.Database("FLAR-pSQL2017\pSQL2017", "RiesgoDB", [Query="SELECT * FROM DWH.METRICAS_PORTAFOLIOS_VIEW WHERE atributo='TRR Index Val LOC'"]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"valor", type number}, {"Fecha", type date}})
in
    #"Changed Type"
