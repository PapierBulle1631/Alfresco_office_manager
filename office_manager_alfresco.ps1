Add-Type -AssemblyName 'System.Windows.Forms'

# En cas de bseoin ces variables donnent la résolution de l'écran
$screen = [System.Windows.Forms.Screen]::PrimaryScreen
$width = $screen.Bounds.Width
$height = $screen.Bounds.Height





#############################################
#                                           #
#    Création des éléments de la fenêtre    #
#                                           #
#############################################

$form = New-Object System.Windows.Forms.Form
$form.Text = "Scanner d'ancienne version office"
$form.Size = New-Object System.Drawing.Size(600, 400)
$form.MinimumSize = New-Object System.Drawing.Size(600, 400)

# Label pour source
$sourceLabel = New-Object System.Windows.Forms.Label
$sourceLabel.Text = 'Dossier de recherche :'
$sourceLabel.Location = New-Object System.Drawing.Point(10, 20)
$sourceLabel.Width = 120
$sourceLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($sourceLabel)

# Zone de texte source
$sourceTextBox = New-Object System.Windows.Forms.TextBox
$sourceTextBox.Location = New-Object System.Drawing.Point(130, 20)
$sourceTextBox.Width = 320
$sourceTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($sourceTextBox)

# Bouton source
$sourceButton = New-Object System.Windows.Forms.Button
$sourceButton.Text = 'Parcourir'
$sourceButton.Location = New-Object System.Drawing.Point(460, 20)
$sourceButton.Width = 100
$sourceButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($sourceButton)

# Labe pour destination
$destinationLabel = New-Object System.Windows.Forms.Label
$destinationLabel.Text = 'Dossier de destination :'
$destinationLabel.Location = New-Object System.Drawing.Point(10, 60)
$destinationLabel.Width = 130
$destinationLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($destinationLabel)

# Zone de texte de destination
$destinationTextBox = New-Object System.Windows.Forms.TextBox
$destinationTextBox.Location = New-Object System.Drawing.Point(140, 60)
$destinationTextBox.Width = 310
$destinationTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($destinationTextBox)

# Bouton de destination
$destinationButton = New-Object System.Windows.Forms.Button
$destinationButton.Text = 'Parcourir'
$destinationButton.Location = New-Object System.Drawing.Point(460, 60)
$destinationButton.Width = 100
$destinationButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($destinationButton)

# Zone de texte des logs
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Location = New-Object System.Drawing.Point(10, 100)
$logTextBox.Width = 550
$logTextBox.Height = 200
$logTextBox.Multiline = $true
$logTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$logTextBox.ReadOnly = $true
$logTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$form.Controls.Add($logTextBox)

# Case à cocher pour la conversion
$conversionCheckbox = New-Object System.Windows.Forms.CheckBox
$conversionCheckbox.Text = "Convertir les fichiers"
$conversionCheckbox.Location = New-Object System.Drawing.Point(70, 315)
$conversionCheckbox.Width = 140
$conversionCheckbox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($conversionCheckbox)

# Bouton de lancement
$processButton = New-Object System.Windows.Forms.Button
$processButton.Text = 'Lancer le programme'
$processButton.Location = New-Object System.Drawing.Point(210, 315)
$graphics = [System.Drawing.Graphics]::FromImage([System.Drawing.Bitmap]::new(1, 1))
$textSize = $graphics.MeasureString($processButton.Text, $processButton.Font)
$processButton.Width = [math]::Ceiling($textSize.Width) + 10 # Ajouter un peu de padding
$processButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($processButton)

# Bouton de suppression
$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = 'Supprimer les fichiers'
$deleteButton.Location = New-Object System.Drawing.Point(350, 315)
$graphics = [System.Drawing.Graphics]::FromImage([System.Drawing.Bitmap]::new(1, 1))
$textSize = $graphics.MeasureString($deleteButton.Text, $deleteButton.Font)
$deleteButton.Width = [math]::Ceiling($textSize.Width) + 10 # Ajouter un peu de padding
$deleteButton.Enabled = $false
$deleteButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($deleteButton)






#########################################################################
#                                                                       #
#   Gestion des fonctions liées aux boutons de recherche des dossiers   #
#                                                                       #
#########################################################################





# Fonction pour choisir le dossier source  des documents à copier
$sourceButton.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Sélectionnez le dossier de recherche. Ce dossier et tout ses enfants seront scannés à la recherche de document avec les extensions .doc, .xls, .ppt puis affichés dans le journal."
    
    if ($folderDialog.ShowDialog() -eq 'OK') {
        $sourceTextBox.Text = $folderDialog.SelectedPath
        $logTextBox.AppendText("Dossier sélectionné pour la source : $($folderDialog.SelectedPath)`r`n")
    } else {
        $logTextBox.AppendText("Aucun dossier source sélectionné.`r`n")
    }
})

# Fonction pour choisir le dossier où sera enregistrée la copie
$destinationButton.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Sélectionnez l'emplacement de la copie."

    if ($folderDialog.ShowDialog() -eq 'OK') {
        $destinationTextBox.Text = $folderDialog.SelectedPath
        $logTextBox.AppendText("Dossier sélectionné pour le rapport : $($folderDialog.SelectedPath)`r`n")
    } else {
        $logTextBox.AppendText("Aucun dossier destination sélectionné.`r`n")
    }
})






################################################################################################
#                                                                                              #
#   Programme principal qui s'éxécute lorsque que le bouton "Lancer le programme" est cliqué   #
#                                                                                              #
################################################################################################







$processButton.Add_Click({
    $sourcePath = $sourceTextBox.Text
    $destinationPath = $destinationTextBox.Text
    $convertFiles = $conversionCheckbox.Checked  # Variable vérifiant l'état de la cas à cocher pour la suppression des fichiers

    # Gestion d'erreur en cas de chemin non valide ou non existant
    if ([String]::IsNullOrWhiteSpace($sourcePath)) {
        $logTextBox.AppendText("Erreur : Le dossier source est vide. Veuillez sélectionner un chemin valide pour continuer`r`n")
        return
    }

    if ([String]::IsNullOrWhiteSpace($destinationPath)) {
        $logTextBox.AppendText("Erreur : Le dossier de destination est vide. Veuillez sélectionner un chemin valide pour continuer`r`n")
        return
    }

    if (-not (Test-Path $sourcePath)) {
        $logTextBox.AppendText("Erreur : Le dossier source spécifié n'existe pas.`r`n")
        return
    }

    if (-not (Test-Path $destinationPath)) {
        $logTextBox.AppendText("Le dossier destination spécifié n'existe pas. Création du dossier...`r`n`r`n")
        try {
            New-Item -Path $destinationPath -ItemType Directory
            $logTextBox.AppendText("Le dossier de destination a bien été créé`r`n")
        }
        catch{
            $logTextBox.AppendText("Erreur : le dossier de destination n'a pas été créé. Veuillez réessayer.")
        }

    }









    # Le runspace est le nom du conteneur de l'éxécution en arrière plan
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()







    ##########################################
    #                                        #
    #   Programme éxécuté en arrière plan    #
    #                                        #
    ##########################################









    # Creation d'un script powershell qui tourne en arrière-plan en même temps
    $runspaceScriptBlock = {
        param($sourcePath, $destinationPath, $logTextBox, $form, $deleteButton, $convertFiles)


        $listeDesFichiers = @()


        # Copier les fichiers en conservant l'arborescence
        Get-ChildItem -Path $sourcePath -Recurse -Include *.doc, *.xls, *.ppt | ForEach-Object {
            $destinationFile = $_.FullName.Replace($sourcePath, $destinationPath)
            $destinationDir = Split-Path -Path $destinationFile -Parent

            if (-Not (Test-Path -Path $destinationDir)) {
                New-Item -Path $destinationDir -ItemType Directory
            }

            Copy-Item -Path $_.FullName -Destination $destinationFile


            #############################################################################
            #                                                                           #
            #   Création du csv avec els données qui nous intéresse (donc à modifier)   #
            #                                                                           #
            #############################################################################


            
            #Version avec uniquement el chemin complet :
            $listeDesFichiers += [PSCustomObject]@{FilePath = $_.FullName}


            # Ajout du chemin | nom du fichiers à la liste qui va servir pour le csv
                # $cheminSeul = ($_.FullName.Substring(0, $_.fullName.Length - $_.Name.Length))
                # $listeDesFichiers += [PSCustomObject]@{Chemin = "$cheminSeul"; Fichier = $_.Name}


            # Version avec uniquement le nom de fichiers :
                # $listeDesFichiers += [PSCustomObject]@{FilePath = $_.Name}


            


            # On affiche l'action en cours : la copie d'un fichier
            $logTextBox.Invoke([Action]{
                    $logTextBox.AppendText("Copie du fichier $($_.FullName).`r`n")
            })
                                
        
        }


        
        $outputCsv = Join-Path -Path $destinationPath -ChildPath "liste_fichiers_e1.xlsx"
        
        $listeDesFichiers | Export-Excel -Path $outputCsv -WorksheetName "Feuille1" -AutoSize
        $logTextBox.Invoke([Action]{
            $logTextBox.AppendText("`r`nCréation de la liste des fichiers E1 : $outputCsv`r`n")
        })



        ##########################################################################################
        #                                                                                        #
        #   Conversion des fichiers du dossier contenant les copies en fonction de l'extension   #
        #                                                                                        #
        ##########################################################################################



        
        
        if ($convertFiles){

            $logTextBox.Invoke([Action]{
                $logTextBox.AppendText("`r`n Délai de synchronisation des premiers fichiers...`r`n")
                Start-Sleep -Seconds 3
                $logTextBox.AppendText("Début de la conversion des fichiers...`r`n")
            })

            $listeFichiersConvertis = @()


            Get-ChildItem -Path $destinationPath -Recurse -File | ForEach-Object {
                # Fonction pour fermer tout processus office qui s'ouvrirait pendant la conversion
                function CleanUp-OfficeProcesses {
                    $officeProcesses = Get-Process | Where-Object { $_.Name -in @("WINWORD", "EXCEL", "POWERPNT") }
                    if ($officeProcesses) {
                        $officeProcesses | ForEach-Object {
                            try {
                                # Force pour tuer les processus Offices
                                Stop-Process -Name $_.Name -Force
                                Write-Host "Arrêt du processus Ofiice pour : $($_.Name)"
                            }
                            catch {
                                Write-Host "Erreur lors de l'arrêt du processus pour : $($_.Name)"
                            }
                        }
                    } else {
                        Write-Host "Aucun processus Office n'est en cours d'éxécution."
                    }
                }



                # Modifie les étapes de conversion en fonction de l'extension
                $tempPath = $_.FullName 
                $tempName = $_.Name
                switch ($_.Extension.ToLower()) {
                    '.doc' {
                        # Convertis les .doc en .docx
                        $newDocxFile = "$($tempPath.Substring(0, $tempPath.Length - 4)).docx"
                        $word = New-Object -ComObject Word.Application
                        $word.Visible = $false
                        $word.DisplayAlerts = $false
                        try {
                            $document = $word.Documents.Open($tempPath)
                            $document.SaveAs($newDocxFile, 12)  # 12 corresponds au format .docx
                            $document.Close()
                            $word.Quit()
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null

                            # Affichage dans les logs
                            $logTextBox.Invoke([Action]{
                                $logTextBox.AppendText("Conversion terminée : $($newDocxFile)`r`n")
                            })
                            $listeFichiersConvertis += [PSCustomObject]@{Nm_Orig = $newDocxFile; Nm_Tmp=$tempName}
                        }
                    
                        catch {
                            # Message d'erreur si la conversion écoue
                            $logTextBox.Invoke([Action]{
                                $logTextBox.AppendText("Erreur lors de la conversion de : $tempPath`r`n")
                            })
                        }
                        finally {
                            # Nettoie toutes les instances ouvertes
                            if ($document) {
                                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
                            }
                            if ($word) {
                                $word.Quit()  # Quit PowerPoint even if there was an error
                                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                            }
                        }
                    }
                    '.xls' {
                        # Convertis de .xls vers .xlsx
                        $newXlsxFile = "$($tempPath.Substring(0, $tempPath.Length - 4)).xlsx"
                        $excel = New-Object -ComObject Excel.Application
                        $excel.Visible = $false
                        $excel.DisplayAlerts = $false
                        try {
                            $workbook = $excel.Workbooks.Open($tempPath)
                            $workbook.SaveAs($newXlsxFile, 51)  # 51 corresponds au format .xlsx
                            $workbook.Close()
                            $excel.Quit()
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null

                            # Affichage dans les logs
                            $logTextBox.Invoke([Action]{
                                $logTextBox.AppendText("Conversion terminée : $($newXlsxFile)`r`n")
                            })
                            $listeFichiersConvertis += [PSCustomObject]@{Nm_Orig = $newXlsxFile; Nm_Tmp=$tempName}
                        }
                        catch {
                            # Message d'erreur si la conversion écoue
                            $logTextBox.Invoke([Action]{
                                $logTextBox.AppendText("Erreur lors de la conversion de : $tempPath`r`n")
                            })
                        }
                        finally {
                            # Nettoie toutes les instances ouvertes
                            if ($workbook) {
                                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                            }
                            if ($excel) {
                                $excel.Quit()  # Quit PowerPoint even if there was an error
                                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                            }
                        }
                    }
                    '.ppt' {
                        # Convertis les .ppt vers .pptx
                        $newPptxFile = "$($tempPath.Substring(0, $tempPath.Length - 4)).pptx"
                        $powerpoint = New-Object -ComObject PowerPoint.Application
                        $powerpoint.Visible = $false        
                        $powerpoint.DisplayAlerts = $false  

                        ###################################################################################################################
                        #                                                                                                                 #
                        # Ici la méthode est différente car la méthode utilisée pour les autres doc a été bloquée par Microsoft pour      #
                        # les Powerpoint. Source : https://learn.microsoft.com/en-us/office/compatibility/office-file-format-reference    #
                        #                                                                                                                 #
                        ###################################################################################################################


                        try {
                            # Ouvre et réenregistre la présentation PowerPoint
                            # Ouvre le powerpoint en mode Compatibilité avec les paramètres (readonly, untitled, password)
                            $presentation = $powerpoint.Presentations.Open($tempPath, $true, $false, $false)


                            # Très important pour PowerPoint : ne pas utilisé le SaveAs avec le code d'extension car cela ne marche plus

                            $presentation.SaveAs($newPptxFile, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLPresentation)
                            $presentation.Close()
                            $logTextBox.Invoke([Action]{
                                $logTextBox.AppendText("Conversion terminée : $($newPptxFile)`r`n")
                            })
                            $listeFichiersConvertis += [PSCustomObject]@{Nm_Orig = $newPptxFile; Nm_Tmp=$tempName}
                        }
                        catch {
                            # Message d'erreur si la conversion écoue
                            $logTextBox.Invoke([Action]{
                                $logTextBox.AppendText("Erreur lors de la conversion de : $tempPath`r`n")
                            })
                        }
                        finally {
                            # Nettoie toutes les instances ouvertes
                            if ($presentation) {
                                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
                            }
                            if ($powerpoint) {
                                $powerpoint.Quit()  # Quit PowerPoint even if there was an error
                                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
                            }
                        }
                    }
                }
                # Ferme tout les processus office ouverts même s'il y a eu des erreurs
                CleanUp-OfficeProcesses
            }

            $logTextBox.Invoke([Action]{
                $logTextBox.AppendText("`r`nFin de la conversion des fichiers.`r`n")
            })
        

            ##############################################################################
            #                                                                            #
            #   Création de liste E2 au format Excel .xlsx et suppression des fichiers   #
            #                                                                            #
            ##############################################################################


            $outputXlsx = Join-Path -Path $destinationPath -ChildPath "liste_fichiers_e2.xlsx"
            $listeFichiersConvertis | Select-Object Nm_Orig, Nm_Tmp | Export-Excel -Path $outputXlsx -WorksheetName "Feuille1" -AutoSize


            $logTextBox.Invoke([Action]{
                $logTextBox.AppendText("`r`nLa liste des fichiers E2 : $outputXlsx a bien été créé.`r`n")
            })


        } # Fin de la phase de conversion


        
        # Active le bouton de suppression maintenant que les fichiers ont été copiés
        $form.Invoke([Action]{
            $deleteButton.Enabled = $true
        })


        


        # Fonction de suppression des fichiers et dossiers présents dans le dossier de copie
        $deleteButton.Add_Click({
            $logTextBox.AppendText("`r`nDébut de la suppression des fichiers avec les anciennes extensions.`r`n")
            Get-ChildItem -Path $destinationPath -Recurse -Include *.doc, *.xls, *.ppt | ForEach-Object {
                try {
                    Remove-Item -Path $_.FullName -Force -Recurse
                    $logTextBox.AppendText("Elément supprimé : $($_.FullName)`r`n")
                } catch {
                    $logTextBox.AppendText("Erreur lors de la suppression de l'élément : $($_.FullName)`r`n")
                }
            }

            $logTextBox.AppendText("`r`nFin de la suppression des fichiers avec les anciennes extensions.`r`n")

        }) #fin de la fonction du bouton delete



    }


    ##########################################
    #                                        #
    #   Fin de l'éxécution en arrière plan   #
    #                                        #
    ##########################################



    # Execute le script en arrière plan au click sur le bouton (et ajoute les bon arguments)
    $runspaceThread = [powershell]::Create().AddScript($runspaceScriptBlock).AddArgument($sourcePath).AddArgument($destinationPath).AddArgument($logTextBox).AddArgument($form).AddArgument($deleteButton).AddArgument($convertFiles)
    $runspaceThread.BeginInvoke()





    # Notification de début du programme en arrière plan
    $logTextBox.AppendText("`r`n`rLe processus de recherche des fichiers a démarré.`r`n`r`n")
})






# Affiche la fenêtre principale
$form.ShowDialog()
