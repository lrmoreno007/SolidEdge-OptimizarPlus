Imports System.Runtime.InteropServices
Imports SolidEdgeFramework

Module Module1
    <STAThread()>
    Sub Main()
        Dim instances As Process() = Process.GetProcessesByName("Edge")
        If instances.Count >= 2 Then
            MsgBox("Más de 1 aplicación de SolidEdge abierta, utilice solo una y vuelva a intentarlo.")
            Return
        End If
        Dim objApplication As SolidEdgeFramework.Application
        objApplication = Marshal.GetActiveObject("SolidEdge.Application")
        'objApplication = cType(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
        'objApplication = DirectCast(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
        'objApplication = TryCast(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
        Dim objDocuments As SolidEdgeFramework.Documents
        objDocuments = objApplication.Documents
        Dim activeDocument As SolidEdgeFramework.SolidEdgeDocument
        If objDocuments.Count > 0 Then
            activeDocument = CType(ObjApplication.ActiveDocument, SolidEdgeFramework.SolidEdgeDocument)
        Else
            MsgBox("Ningún documento abierto")
            Console.WriteLine("Ningún documento abierto")
            Return
        End If
        Dim objModels As SolidEdgePart.Models = Nothing
        Console.Write("Documento del tipo: ")
        Try
            Select Case activeDocument.Type
                Case SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument
                    MsgBox("Ensamblaje, nada que hacer")
                    Console.WriteLine("Ensamblaje")
                    objApplication.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewFit)
                    Return
                Case SolidEdgeFramework.DocumentTypeConstants.igDraftDocument
                    MsgBox("Plano, nada que hacer")
                    Console.WriteLine("Plano")
                    objApplication.StartCommand(SolidEdgeConstants.DetailCommandConstants.DetailViewFit)
                    Return
                Case SolidEdgeFramework.DocumentTypeConstants.igPartDocument
                    Console.WriteLine("Pieza")
                    Dim objPartDocument As SolidEdgePart.PartDocument
                    objPartDocument = objApplication.ActiveDocument
                    objModels = objPartDocument.Models
                    objPartDocument.CoordinateSystems.Visible = False
                    objPartDocument.RefPlanes.Item(1).Visible = False
                    objPartDocument.RefPlanes.Item(2).Visible = False
                    objPartDocument.RefPlanes.Item(3).Visible = False
                    objApplication.StartCommand(SolidEdgeConstants.PartCommandConstants.PartViewFit)
                    Try
                        Console.Write("Cambiando a Modo Síncrono... ")
                        If CType(SolidEdgePart.ModelingModeConstants.seModelingModeOrdered, Boolean) Then
                            objPartDocument.ModelingMode =
                                CType(SolidEdgePart.ModelingModeConstants.seModelingModeSynchronous,
                                      SolidEdgePart.ModelingModeConstants)
                        End If
                        Console.WriteLine("OK")
                    Catch
                        Console.WriteLine("ERROR: El modelo está en Modo Ordenado")
                        Console.Write("Moviendo el modelo a Modo Síncrono... ")
                        Dim objFeatures As SolidEdgePart.Features = Nothing
                        Dim objFeature As Object = Nothing
                        Dim objModelPart As SolidEdgePart.Model
                        objModelPart = objModels.Item(1)
                        objFeatures = objModelPart.Features
                        Dim bIgnoreWarnings As Boolean = True
                        Dim bExtentSelection As Boolean = True
                        Dim aErrorMessages As Array
                        Dim aWarningMessages As Array
                        Dim lNumberOfFeaturesCausingError As Integer
                        Dim lNumberOfFeaturesCausingWarning As Integer
                        For Each objFeature In objFeatures
                            aErrorMessages = Array.CreateInstance(GetType(String), 0)
                            aWarningMessages = Array.CreateInstance(GetType(String), 0)
                            Dim dVolumeDifference As Double = 0
                            'MoveToSynchronous en pieza tiene 8 argumentos
                            objPartDocument.MoveToSynchronous(objFeature, bIgnoreWarnings, bExtentSelection,
                                                              lNumberOfFeaturesCausingError, aErrorMessages,
                                                              lNumberOfFeaturesCausingWarning, aWarningMessages,
                                                              dVolumeDifference)
                        Next
                        objPartDocument.ModelingMode =
                            CType(SolidEdgePart.ModelingModeConstants.seModelingModeSynchronous,
                                  SolidEdgePart.ModelingModeConstants)
                        Console.WriteLine("OK")
                    Finally
                    End Try

                Case SolidEdgeFramework.DocumentTypeConstants.igSheetMetalDocument
                    Console.WriteLine("Chapa")
                    Dim objPartDocument As SolidEdgePart.SheetMetalDocument
                    objPartDocument = objApplication.ActiveDocument
                    objModels = objPartDocument.Models
                    objPartDocument.CoordinateSystems.Visible = False
                    objPartDocument.RefPlanes.Item(1).Visible = False
                    objPartDocument.RefPlanes.Item(2).Visible = False
                    objPartDocument.RefPlanes.Item(3).Visible = False
                    objApplication.StartCommand(SolidEdgeConstants.SheetMetalCommandConstants.SheetMetalViewFit)
                    Try
                        Console.Write("Cambiando a Modo Síncrono... ")
                        If CType(SolidEdgePart.ModelingModeConstants.seModelingModeOrdered, Boolean) Then
                            objPartDocument.ModelingMode =
                                CType(SolidEdgePart.ModelingModeConstants.seModelingModeSynchronous,
                                      SolidEdgePart.ModelingModeConstants)
                        End If
                        Console.WriteLine("OK")
                    Catch
                        Console.WriteLine("ERROR: El modelo está en Modo Ordenado")
                        Console.Write("Moviendo el modelo a Modo Síncrono... ")
                        Dim objFeatures As SolidEdgePart.Features = Nothing
                        Dim objFeature As Object = Nothing
                        Dim objModelSheetMetal As SolidEdgePart.Model
                        objModelSheetMetal = objModels.Item(1)
                        objFeatures = objModelSheetMetal.Features
                        Dim bIgnoreWarnings As Boolean = True
                        Dim bExtentSelection As Boolean = True
                        Dim aErrorMessages As Array
                        Dim aWarningMessages As Array
                        Dim lNumberOfFeaturesCausingError As Integer
                        Dim lNumberOfFeaturesCausingWarning As Integer
                        For Each objFeature In objFeatures
                            aErrorMessages = Array.CreateInstance(GetType(String), 0)
                            aWarningMessages = Array.CreateInstance(GetType(String), 0)
                            'MoveToSynchronous en chapa tiene 7 argumentos
                            objPartDocument.MoveToSynchronous(objFeature, bIgnoreWarnings, bExtentSelection,
                                                              lNumberOfFeaturesCausingError, aErrorMessages,
                                                              lNumberOfFeaturesCausingWarning, aWarningMessages)
                        Next
                        objPartDocument.ModelingMode =
                            CType(SolidEdgePart.ModelingModeConstants.seModelingModeSynchronous,
                                  SolidEdgePart.ModelingModeConstants)
                        Console.WriteLine("OK")
                    Finally
                    End Try

                Case SolidEdgeFramework.DocumentTypeConstants.igWeldmentAssemblyDocument
                    MsgBox("Ensamblaje soldado, nada que hacer")
                    Console.WriteLine("Ensamblaje soldado")
                    objApplication.StartCommand(SolidEdgeConstants.WeldmentCommandConstants.WeldmentViewFit)
                    Return
                Case SolidEdgeFramework.DocumentTypeConstants.igWeldmentDocument
                    MsgBox("Soldadura, nada que hacer")
                    Console.WriteLine("Soldadura")
                    objApplication.StartCommand(SolidEdgeConstants.WeldmentCommandConstants.WeldmentViewFit)
                    Return
            End Select
            ' OPTIMIZACIÓN
            Console.Write("Optimizando el modelo... ")
            Dim objModel As SolidEdgePart.Model
            objModel = objModels.Item(1)
            objModel.HealAndOptimizeBody(False, True)
            objApplication.DoIdle()
            Console.WriteLine("OK")
            ' RECONOCER AGUJEROS
            Console.Write("Reconociendo agujeros... ")
            objModel = objModels.Item(1)
            Dim numBodies As Integer = 1
            Dim objModelBody As SolidEdgeGeometry.Body
            objModelBody = CType(objModel.Body, SolidEdgeGeometry.Body)
            Dim objBodies As Array
            objBodies = New SolidEdgeGeometry.Body(0) {objModelBody}
            Dim numHoles As Integer = 1
            Dim objRecognizedHoles As Array
            objRecognizedHoles = New SolidEdgePart.Features() {}
            objModel.Holes.RecognizeAndCreateHoleGroups(numBodies, objBodies, numHoles, objRecognizedHoles)
            objApplication.DoIdle()
            Console.WriteLine("OK")
            ' CAMBIAR A MODO ORDENADO
            Console.Write("Cambiando a Modo Ordenado... ")
            objModel.Recompute()
            Select Case activeDocument.Type
                Case SolidEdgeFramework.DocumentTypeConstants.igPartDocument
                    Dim objPartDocument As SolidEdgePart.PartDocument
                    objPartDocument = CType(objApplication.ActiveDocument, SolidEdgePart.PartDocument)
                    If CType(SolidEdgePart.ModelingModeConstants.seModelingModeSynchronous, Boolean) Then
                        objPartDocument.ModelingMode = CType(SolidEdgePart.ModelingModeConstants.seModelingModeOrdered,
                                                             SolidEdgePart.ModelingModeConstants)
                    End If
                    Console.WriteLine("OK")

                Case SolidEdgeFramework.DocumentTypeConstants.igSheetMetalDocument
                    Dim objPartDocument As SolidEdgePart.SheetMetalDocument
                    objPartDocument = CType(objApplication.ActiveDocument, SolidEdgePart.SheetMetalDocument)
                    If CType(SolidEdgePart.ModelingModeConstants.seModelingModeSynchronous, Boolean) Then
                        objPartDocument.ModelingMode = CType(SolidEdgePart.ModelingModeConstants.seModelingModeOrdered,
                                                             SolidEdgePart.ModelingModeConstants)
                    End If
                    Console.WriteLine("OK")

                    Console.Write("Buscando la cara mas grande... ")
                    If objModels.Count > 0 Then
                        Dim objBody As SolidEdgeGeometry.Body = CType(objModel.Body, SolidEdgeGeometry.Body)
                        Dim objFaces As SolidEdgeGeometry.Faces =
                                objBody.Faces(SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryAll)
                        objApplication.DoIdle()
                        Dim objBaseFace As SolidEdgeGeometry.Face = objFaces.Item(1)
                        Dim maxArea As Double = 0
                        For i As Integer = 1 To objFaces.Count
                            Dim f As SolidEdgeGeometry.Face = objFaces.Item(i)
                            If f.Area > maxArea Then
                                maxArea = f.Area
                                objBaseFace = f
                            End If
                        Next
                        Console.WriteLine("OK")

                        Console.Write("Contando el número de aristas de la pieza...  ")
                        Dim objAllEdges As SolidEdgeGeometry.Edges =
                                objBody.Edges(SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryAll)
                        Dim aEdges As Array = Array.CreateInstance(GetType(Object), objAllEdges.Count)
                        For i As Integer = 1 To objAllEdges.Count
                            aEdges.SetValue(objAllEdges.Item(i), i - 1)
                        Next
                        Console.Write("OK ---> Encontradas ")
                        Console.Write(objAllEdges.Count)
                        Console.WriteLine(" aristas en la pieza")
                        Console.Write("Transformando a chapa ")
                        Try
                            objModel.ConvToSMs.AddEx(objBaseFace, 1, aEdges, Nothing, 0.0, 0.0)
                            Console.WriteLine("OK")
                        Catch ex As Exception
                            Console.WriteLine("Error: " & ex.Message)
                        End Try
                    End If
            End Select

            Dim respuestaGuardar As MsgBoxResult = MsgBox("¿Desea guardar el documento?",
                                                          MsgBoxStyle.YesNo + MsgBoxStyle.Question, 
                                                          "Guardar Como")
            If respuestaGuardar = MsgBoxResult.Yes Then
                Try
                    objApplication.DoIdle()
                    activeDocument.Save()
                Catch ex As Exception
                    Console.WriteLine("Error al guardar: " & ex.Message)
                    MsgBox("No se pudo realizar el Guardar Como. Compruebe si el archivo ya existe o está abierto.",
                           MsgBoxStyle.Critical)
                End Try
            End If

            ' La selección del material tiene que ser la última porque este formulario no detiene la ejecución de la macro y finaliza.
            Dim respuesta As MsgBoxResult = MsgBox("¿Desea asignar un material a la pieza?",
                                                   MsgBoxStyle.YesNo + MsgBoxStyle.Question,
                                                   "Asignar Material")
            If respuesta = MsgBoxResult.Yes Then
                Try
                    objApplication.DoIdle()
                    objApplication.StartCommand(45163)
                    objApplication.DoIdle()
                Catch ex As Exception
                    Console.WriteLine("Error al abrir la tabla de materiales: " & ex.Message)
                End Try
            End If

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            'MsgBox(ex.Message)
            System.Environment.ExitCode = - 1
        Finally
            If System.Environment.ExitCode = 0 Then
                Console.WriteLine("Finalizado OK")
            Else
                Console.WriteLine("Finalizado con errores")
            End If
        End Try
    End Sub
End Module
