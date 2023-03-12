Begin Dialog NuevoDiálogo 50,50,150,150,"NuevoDiálogo", .NuevoDiálogo
End Dialog
'Desarrollado por Jorge M. Chávez
'Fecha: 01/03/2023

Sub Main
	IgnoreWarning(True)
	Call ReportReaderImport1()	'D:\RUC1\DATA\Archivos fuente.ILB\PDT_621_2022.pdf
	Call ReportReaderImport2()	'D:\RUC1\DATA\Archivos fuente.ILB\PDT_621_2022.pdf
	Call ReportReaderImport3()	'D:\RUC1\DATA\Archivos fuente.ILB\PDT_621_2022.pdf
	Call ReportReaderImport4()	'D:\RUC1\DATA\Archivos fuente.ILB\PDT_621_2022.pdf
	Call AppendDatabase()		'PDT621_TRIBUTO1.IMD
	Call AppendField()		'PDT621.IMD
	Call Summarization1()		'PDT621_base.IMD
	Call Summarization2()		'PDT621_base.IMD
	Call ExportDatabaseXLSX()	'PDT621.IMD
	Client.CloseAll
	Client.DeleteDatabase "PDT621_BASE1.IMD"
	Client.DeleteDatabase "PDT621_TRIBUTO1.IMD"
	Client.DeleteDatabase "PDT621_DDRTA.IMD"
	Client.DeleteDatabase "PDT621_DDIGV.IMD"
	Dim pm As Object
	Dim SourcePath As String
	Dim DestinationPath As String
	Set SourcePath = Client.WorkingDirectory
	Set DestinationPath = "D:\RUC1\DATA\_PDT621"
	Client.RunAtServer False
	Set pm = Client.ProjectManagement
	pm.MoveDatabase SourcePath + "A_PDT621.IMD", DestinationPath
	pm.MoveDatabase SourcePath + "A.1_PDT621_Rta_mensual.IMD", DestinationPath
	pm.MoveDatabase SourcePath + "A.2_PDT621_IGV_mensual.IMD", DestinationPath
	Set pm = Nothing
	Client.RefreshFileExplorer
End Sub


' Archivo - Asistente de importación: Report Reader
Function ReportReaderImport1
	dbName = "PDT621_BASE1.IMD"
	Client.ImportPrintReportEx "D:\RUC1\DATA\Definiciones de importación.ILB\PDT_621_base.jpm", "D:\RUC1\DATA\Archivos fuente.ILB\PDT_621_2022.pdf", dbname, FALSE, FALSE
End Function

' Archivo - Asistente de importación: Report Reader
Function ReportReaderImport2
	dbName = "PDT621_TRIBUTO1.IMD"
	Client.ImportPrintReportEx "D:\RUC1\DATA\Definiciones de importación.ILB\PDT_621_tributo.jpm", "D:\RUC1\DATA\Archivos fuente.ILB\PDT_621_2022.pdf", dbname, FALSE, FALSE
End Function


' Archivo - Asistente de importación: Report Reader
Function ReportReaderImport3
	dbName = "PDT621_DDRTA.IMD"
	Client.ImportPrintReportEx "D:\RUC1\DATA\Definiciones de importación.ILB\PDT_621_dd_rta.jpm", "D:\RUC1\DATA\Archivos fuente.ILB\PDT_621_2022.pdf", dbname, FALSE, FALSE
End Function


' Archivo - Asistente de importación: Report Reader
Function ReportReaderImport4
	dbName = "PDT621_DDIGV.IMD"
	Client.ImportPrintReportEx "D:\RUC1\DATA\Definiciones de importación.ILB\PDT_621_dd_igv.jpm", "D:\RUC1\DATA\Archivos fuente.ILB\PDT_621_2022.pdf", dbname, FALSE, FALSE
End Function

' Archivo: Anexar bases de datos
Function AppendDatabase
	Set db = Client.OpenDatabase("PDT621_BASE1.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "PDT621_TRIBUTO1.IMD"
	task.AddDatabase "PDT621_DDIGV.IMD"
	task.AddDatabase "PDT621_DDRTA.IMD"
	dbName = "A_PDT621.IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
End Function

' Anexar campo
Function AppendField
	Set db = Client.OpenDatabase("A_PDT621.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "BASE_PDT_IGV"
	field.Description = ""
	field.Type = WI_VIRT_NUM
	field.Equation = "@if(@match( CAS ; ""100""; ""160"");  MONTO ;@If(@match( CAS ; ""102""; ""162""); - MONTO;  0.00))"
	field.Decimals = 2
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Análisis: Resumen
Function Summarization1
	Set db = Client.OpenDatabase("A_PDT621.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "PERIODO"
	task.AddFieldToTotal "MONTO"
	task.Criteria = " CAS  = ""301"""
	dbName = "A.1_PDT621_Rta_mensual.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Análisis: Resumen
Function Summarization2
	Set db = Client.OpenDatabase("A_PDT621.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "PERIODO"
	task.AddFieldToTotal "BASE_PDT_IGV"
	dbName = "A.2_PDT621_IGV_mensual.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Archivo-Exportar base de datos: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("A_PDT621.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "D:\RUC1\DATA\Exportaciones.ILB\A_PDT621.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function