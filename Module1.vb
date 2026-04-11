Imports System.Runtime.InteropServices
Imports SolidEdgeFramework

Module Module1
    <STAThread()>
    Sub Main()
        Dim instances As Process() = Process.GetProcessesByName("Edge") 'Obtener instancias de procesos con nombre "Edge"
        If instances.Count >= 2 Then 'Si hay al menos 2 instancias, mostrar mensaje y salir
            MsgBox("Más de 1 aplicación de SolidEdge abierta, utilice solo una y vuelva a intentarlo.")
            Return
        End If
        Dim objApplication As SolidEdgeFramework.Application
        objApplication = Marshal.GetActiveObject("SolidEdge.Application") 'Obtener objeto de la aplicación
        'objApplication = cType(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
        'objApplication = DirectCast(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
        'objApplication = TryCast(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
        Dim objDocuments As SolidEdgeFramework.Documents
        objDocuments = objApplication.Documents 'Obtener documentos de la aplicación
        Dim activeDocument As SolidEdgeFramework.SolidEdgeDocument
        If objDocuments.Count > 0 Then 'Si hay al menos un documento abierto
            activeDocument = CType(ObjApplication.ActiveDocument, SolidEdgeFramework.SolidEdgeDocument) 'Convertir el documento activo a tipo específico
        Else
            MsgBox("Ningún documento abierto")
            Console.WriteLine("Ningún documento abierto")
            Return
        End If
        Dim objModels As SolidEdgePart.Models = Nothing
        Console.Write("Documento del tipo: ")
        Try
            Select Case activeDocument.Type 'Seleccionar el caso según el tipo de documento
                Case SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument 'Si es un ensamblaje
                    MsgBox("Ensamblaje, nada que hacer")
                    Console.WriteLine("Ensamblaje")
                    objApplication.StartCommand(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewFit) 'Ajustar vista del ensamblaje
                    Return
                Case SolidEdgeFramework.DocumentTypeConstants.igDraftDocument 'Si es un plano
                    MsgBox("Plano, nada que hacer")
                    Console.WriteLine("Plano")
                    objApplication.StartCommand(SolidEdgeConstants.DetailCommandConstants.DetailViewFit) 'Ajustar vista del plano
                    Return
                Case SolidEdgeFramework.DocumentTypeConstants.igPartDocument 'Si es una pieza
                    Console.WriteLine("Pieza")
                    Dim objPartDocument As SolidEdgePart.PartDocument
                    objPartDocument = objApplication.ActiveDocument 'Obtener documento activo como parte
                    objModels = objPartDocument.Models 'Obtener modelos de la parte
                    objPartDocument.CoordinateSystems.Visible = False 'Ocultar sistemas de coordenadas
                    objPartDocument.RefPlanes.Item(1).Visible = False 'Ocultar planos de referencia
                    objPartDocument.RefPlanes.Item(2).Visible = False 
                    objPartDocument.RefPlanes.Item(3).Visible = False 
                    objApplication.StartCommand(SolidEdgeConstants.PartCommandConstants.PartViewFit) 'Ajustar vista de la parte
                    Try
                        Console.Write("Cambiando a Modo Síncrono... ")
                        If CType(SolidEdgePart.ModelingModeConstants.seModelingModeOrdered, Boolean) Then 'Si está en modo ordenado
                            objPartDocument.ModelingMode = 
                                CType(SolidEdgePart.ModelingModeConstants.seModelingModeSynchronous,
                                      SolidEdgePart.ModelingModeConstants)
                        End If
                        Console.WriteLine("OK")
                    Catch
                        Console.WriteLine("ERROR: El modelo está en Modo Ordenado")
                        Console.Write("Moviendo el modelo a Modo Síncrono... ")
                        Dim objFeatures As SolidEdgePart.Features = Nothing 'Inicializar características
                        Dim objFeature As Object = Nothing
                        Dim objModelPart As SolidEdgePart.Model
                        objModelPart = objModels.Item(1) 'Obtener primer modelo
                        objFeatures = objModelPart.Features 'Obtener características del modelo
                        Dim bIgnoreWarnings As Boolean = True 'Ignorar advertencias
                        Dim bExtentSelection As Boolean = True 'Seleccionar extensión
                        Dim aErrorMessages As Array
                        Dim aWarningMessages As Array
                        Dim lNumberOfFeaturesCausingError As Integer
                        Dim lNumberOfFeaturesCausingWarning As Integer
                        Dim dVolumeDifference As Integer
                        For Each objFeature In objFeatures 'Iterar sobre características
                            aErrorMessages = Array.CreateInstance(GetType(String), 0) 'Inicializar mensajes de error
                            aWarningMessages = Array.CreateInstance(GetType(String), 0) 'Inicializar mensajes de advertencia
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

                Case SolidEdgeFramework.DocumentTypeConstants.igSheetMetalDocument 'Si es una chapa
                    Console.WriteLine("Chapa")
                    Dim objPartDocument As SolidEdgePart.SheetMetalDocument
                    objPartDocument = objApplication.ActiveDocument 'Obtener documento activo como chapa
                    objModels = objPartDocument.Models 'Obtener modelos de la chapa
                    objPartDocument.CoordinateSystems.Visible = False 'Ocultar sistemas de coordenadas
                    objPartDocument.RefPlanes.Item(1).Visible = False 
                    objPartDocument.RefPlanes.Item(2).Visible = False 
                    objPartDocument.RefPlanes.Item(3).Visible = False 
                    objApplication.StartCommand(SolidEdgeConstants.SheetMetalCommandConstants.SheetMetalViewFit) 'Ajustar vista de la chapa
                    Try
                        Console.Write("Cambiando a Modo Síncrono... ")
                        If CType(SolidEdgePart.ModelingModeConstants.seModelingModeOrdered, Boolean) Then 'Si está en modo ordenado
                            objPartDocument.ModelingMode = 
                                CType(SolidEdgePart.ModelingModeConstants.seModelingModeSynchronous,
                                      SolidEdgePart.ModelingModeConstants)
                        End If
                        Console.WriteLine("OK")
                    Catch
                        Console.WriteLine("ERROR: El modelo está en Modo Ordenado")
                        Console.Write("Moviendo el modelo a Modo Síncrono... ")
                        Dim objFeatures As SolidEdgePart.Features = Nothing 'Inicializar características
                        Dim objFeature As Object = Nothing
                        Dim objModelSheetMetal As SolidEdgePart.Model
                        objModelSheetMetal = objModels.Item(1) 'Obtener primer modelo
                        objFeatures = objModelSheetMetal.Features 'Obtener características del modelo
                        Dim bIgnoreWarnings As Boolean = True 'Ignorar advertencias
                        Dim bExtentSelection As Boolean = True 'Seleccionar extensión
                        Dim aErrorMessages As Array
                        Dim aWarningMessages As Array
                        Dim lNumberOfFeaturesCausingError As Integer
                        Dim lNumberOfFeaturesCausingWarning As Integer
                        For Each objFeature In objFeatures 'Iterar sobre características
                            aErrorMessages = Array.CreateInstance(GetType(String), 0) 'Inicializar mensajes de error
                            aWarningMessages = Array.CreateInstance(GetType(String), 0) 'Inicializar mensajes de advertencia
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

                Case SolidEdgeFramework.DocumentTypeConstants.igWeldmentAssemblyDocument 'Si es un ensamblaje soldado
                    MsgBox("Ensamblaje soldado, nada que hacer")
                    Console.WriteLine("Ensamblaje soldador")
                    objApplication.StartCommand(SolidEdgeConstants.WeldmentCommandConstants.WeldmentViewFit) 'Ajustar vista del ensamblaje soldado
                    Return
                Case SolidEdgeFramework.DocumentTypeConstants.igWeldmentDocument 'Si es una soldadura
                    MsgBox("Soldadura, nada que hacer")
                    Console.WriteLine("Soldadura")
                    objApplication.StartCommand(SolidEdgeConstants.WeldmentCommandConstants.WeldmentViewFit) 'Ajustar vista de la soldadura
                    Return
            End Select
            ' OPTIMIZACIÓN
            Console.Write("Optimizando el modelo... ")
            Dim objModel As SolidEdgePart.Model
            objModel = objModels.Item(1) 'Obtener primer modelo
            objModel.HealAndOptimizeBody(False, True) 'Optimizar cuerpo del modelo
            objApplication.DoIdle() 'Hacer que la aplicación se actualice
            Console.WriteLine("OK")
            ' RECONOCER AGUJEROS
            Console.Write("Reconociendo agujeros... ")
            objModel = objModels.Item(1) 'Obtener primer modelo
            Dim numBodies As Integer = 1 'Número de cuerpos a considerar
            Dim objModelBody As SolidEdgeGeometry.Body
            objModelBody = CType(objModel.Body, SolidEdgeGeometry.Body) 'Convertir cuerpo del modelo
            Dim objBodies As Array
            objBodies = New SolidEdgeGeometry.Body(0) {objModelBody} 'Crear array de cuerpos
            Dim numHoles As Integer = 1 'Número de agujeros a reconocer
            Dim objRecognizedHoles As Array
            objRecognizedHoles = New SolidEdgePart.Features() {} 'Inicializar características de agujeros reconocidos
            objModel.Holes.RecognizeAndCreateHoleGroups(numBodies, objBodies, numHoles, objRecognizedHoles) 'Reconocer y crear grupos de agujeros
            objApplication.DoIdle()
            Console.WriteLine("OK")
            ' CAMBIAR A MODO ORDENADO
            Console.Write("Cambiando a Modo Ordenado... ")
            objModel.Recompute() 'Recalcular modelo
            Select Case activeDocument.Type 'Seleccionar el caso según tipo de documento
                Case SolidEdgeFramework.DocumentTypeConstants.igPartDocument 'Si es una parte
                    Dim objPartDocument As SolidEdgePart.PartDocument
                    objPartDocument = CType(objApplication.ActiveDocument, SolidEdgePart.PartDocument) 'Convertir documento activo a parte
                    If CType(SolidEdgePart.ModelingModeConstants.seModelingModeSynchronous, Boolean) Then 'Si está en modo sincrono
                        objPartDocument.ModelingMode = 
                            CType(SolidEdgePart.ModelingModeConstants.seModelingModeOrdered,
                                  SolidEdgePart.ModelingModeConstants)
                    End If
                    Console.WriteLine("OK")

                Case SolidEdgeFramework.DocumentTypeConstants.igSheetMetalDocument 'Si es una chapa
                    Dim objPartDocument As SolidEdgePart.SheetMetalDocument
                    objPartDocument = CType(objApplication.ActiveDocument, SolidEdgePart.SheetMetalDocument) 'Convertir documento activo a chapa
                    If CType(SolidEdgePart.ModelingModeConstants.seModelingModeSynchronous, Boolean) Then 'Si está en modo sincrono
                        objPartDocument.ModelingMode = 
                            CType(SolidEdgePart.ModelingModeConstants.seModelingModeOrdered,
                                  SolidEdgePart.ModelingModeConstants)
                    End If
                    Console.WriteLine("OK")

                    Console.Write("Buscando la cara mas grande... ")
                    If objModels.Count > 0 Then 'Si hay al menos un modelo
                        Dim objBody As SolidEdgeGeometry.Body = CType(objModel.Body, SolidEdgeGeometry.Body) 'Convertir cuerpo del modelo
                        Dim objFaces As SolidEdgeGeometry.Faces =
                                objBody.Faces(SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryAll) 'Obtener todas las caras
                        objApplication.DoIdle()
                        Dim objBaseFace As SolidEdgeGeometry.Face = objFaces.Item(1) 'Obtener primera cara como base
                        Dim maxArea As Double = 0 'Inicializar área máxima
                        For i As Integer = 1 To objFaces.Count 'Iterar sobre caras
                            Dim f As SolidEdgeGeometry.Face = objFaces.Item(i)
                            If f.Area > maxArea Then 'Si el área es mayor que la actual
                                maxArea = f.Area 'Actualizar área máxima
                                objBaseFace = f 'Actualizar cara base
                            End If
                        Next
                        Console.WriteLine("OK")

                        Console.Write("Contando el número de aristas de la pieza...  ")
                        Dim objAllEdges As SolidEdgeGeometry.Edges =
                                objBody.Edges(SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryAll) 'Obtener todas las aristas
                        Dim aEdges As Array = Array.CreateInstance(GetType(Object), objAllEdges.Count) 'Crear array de aristas
                        For i As Integer = 1 To objAllEdges.Count 'Iterar sobre aristas
                            aEdges.SetValue(objAllEdges.Item(i), i - 1)
                        Next
                        Console.Write("OK ---> Encontradas ")
                        Console.Write(objAllEdges.Count)
                        Console.WriteLine(" aristas en la pieza")
                        Console.Write("Transformando a chapa ")
                        Try
                            objModel.ConvToSMs.AddEx(objBaseFace, 1, aEdges, Nothing, 0.0, 0.0) 'Transformar a chapa
                            Console.WriteLine("OK")
                        Catch ex As Exception
                            Console.WriteLine("Error: " & ex.Message)
                        End Try
                    End If
            End Select

            Dim respuestaGuardar As MsgBoxResult = MsgBox("¿Desea guardar el documento?",
                                                          MsgBoxStyle.YesNo + MsgBoxStyle.Question, 
                                                          "Guardar Como") 'Preguntar si guardar documento
            If respuestaGuardar = MsgBoxResult.Yes Then
                Try
                    objApplication.DoIdle() 'Hacer que la aplicación se actualice
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

        Catch ex As Exception 'Manejar excepciones
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
