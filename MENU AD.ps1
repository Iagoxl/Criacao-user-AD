# === Imports ===

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module ActiveDirectory

function Abrir-CriacaoUsuario {


    Import-Module ActiveDirectory

    # Verifica se o script está em uma política restrita e ajusta temporariamente

    $currentPolicy = Get-ExecutionPolicy -Scope Process
    if ($currentPolicy -ne "Bypass") {
        try {
            Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
            Write-Host "Política de execução temporária ajustada para 'Bypass'." -ForegroundColor Green
        } catch {
            Write-Host "Erro ao ajustar a política de execução. Tente rodar como administrador ou manualmente executar:" -ForegroundColor Red
            Write-Host "Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass" -ForegroundColor Yellow
            exit
        }
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Criação de Usuário AD"
    $form.Size = New-Object System.Drawing.Size(390, 790)
    $form.StartPosition = "CenterScreen"

    $labels = @("Nome Completo", "Login", "Senha", "E-mail", "Departamento", "Cargo")
    $textBoxes = @{}

    for ($i = 0; $i -lt $labels.Length; $i++) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $labels[$i]
        $label.Location = [System.Drawing.Point]::new(10, 20 + ($i * 70))
        $label.Size = [System.Drawing.Size]::new(200, 20)
        $form.Controls.Add($label)

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = [System.Drawing.Point]::new(10, 40 + ($i * 70))
        $textBox.Size = [System.Drawing.Size]::new(350, 20)

        if ($labels[$i] -eq "Senha") {
            $textBox.UseSystemPasswordChar = $true
        }

        $textBoxes[$labels[$i]] = $textBox
        $form.Controls.Add($textBox)
    }

    # Sub-OU como ComboBox
    $labelSubOU = New-Object System.Windows.Forms.Label
    $labelSubOU.Text = "Sub-OU"
    $labelSubOU.Location = [System.Drawing.Point]::new(10, 440)
    $labelSubOU.Size = [System.Drawing.Size]::new(200, 20)
    $form.Controls.Add($labelSubOU)

    $comboBoxSubOU = New-Object System.Windows.Forms.ComboBox
    $comboBoxSubOU.Location = [System.Drawing.Point]::new(10, 460)
    $comboBoxSubOU.Size = [System.Drawing.Size]::new(350, 20)

    # Departamentos ( OU )
    $comboBoxSubOU.DropDownStyle = 'DropDown'
    $comboBoxSubOU.Items.Add("OU=ADICIONA SUA OU")
    $comboBoxSubOU.Items.Add("OU=ADICIONA SUA OU 2")
    

    $comboBoxSubOU.Text = "Adiciona o departamento"
    $form.Controls.Add($comboBoxSubOU)

    # ComboBox para escolher sufixo UPN
    $comboBoxUPNSuffix = New-Object System.Windows.Forms.ComboBox
    $comboBoxUPNSuffix.Location = [System.Drawing.Point]::new(10, 630)
    $comboBoxUPNSuffix.Size = [System.Drawing.Size]::new(350, 20)
    $comboBoxUPNSuffix.DropDownStyle = 'DropDownList'
    $comboBoxUPNSuffix.Items.Add(ADICIONA SEU SUFFIX UPN "@DOMINIO.COM.BR)
    $comboBoxUPNSuffix.SelectedIndex = 0
    $form.Controls.Add($comboBoxUPNSuffix)

    # OU principal
    $labelOU = New-Object System.Windows.Forms.Label
    $labelOU.Text = "OU Principal"
    $labelOU.Location = [System.Drawing.Point]::new(10, 500)
    $labelOU.Size = [System.Drawing.Size]::new(200, 20)
    $form.Controls.Add($labelOU)

    $comboBoxOU = New-Object System.Windows.Forms.ComboBox
    $comboBoxOU.Location = [System.Drawing.Point]::new(10, 520)
    $comboBoxOU.Size = [System.Drawing.Size]::new(350, 20)
    $comboBoxOU.DropDownStyle = 'DropDownList'
    $comboBoxOU.Items.Add("OU=OU PRINCIPAL,DC=DOMINIO,DC=.com.br ou .LOCAL)
    $comboBoxOU.SelectedIndex = 0
    $form.Controls.Add($comboBoxOU)

    # Grupo (opcional)
    $labelGrupo = New-Object System.Windows.Forms.Label
    $labelGrupo.Text = "Grupo (opcional)"
    $labelGrupo.Location = [System.Drawing.Point]::new(10, 560)
    $labelGrupo.Size = [System.Drawing.Size]::new(200, 20)
    $form.Controls.Add($labelGrupo)

    $comboBoxGrupo = New-Object System.Windows.Forms.ComboBox
    $comboBoxGrupo.Location = [System.Drawing.Point]::new(10, 580)
    $comboBoxGrupo.Size = [System.Drawing.Size]::new(350, 20)
    $comboBoxGrupo.DropDownStyle = 'DropDown'

    # Grupos de usuario
    $comboBoxGrupo.Items.Add("ADICIONA SEU GRUPO,ADICIONA SEU GRUPO 2")
    $comboBoxGrupo.Items.Add("ADICIONA SEU GRUPO")
    $form.Controls.Add($comboBoxGrupo)
    $comboBoxGrupo.Text = "Adiciona o grupo"

    # Botão de criação
    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Criar Usuário"
    $button.Location = [System.Drawing.Point]::new(10, 690)
    $button.Size = [System.Drawing.Size]::new(350, 30)
    $button.Add_Click({
        $nome = $textBoxes["Nome Completo"].Text
        $sam = $textBoxes["Login"].Text
        $senha = ConvertTo-SecureString $textBoxes["Senha"].Text -AsPlainText -Force
        $email = $textBoxes["E-mail"].Text
        $departamento = $textBoxes["Departamento"].Text
        $cargo = $textBoxes["Cargo"].Text
        $grupo = $comboBoxGrupo.Text
        $subOU = $comboBoxSubOU.Text
        $ouPrincipal = $comboBoxOU.SelectedItem.ToString()
        $pathOU = "$subOU,$ouPrincipal"

        $partesNome = $nome.Split(" ")
        $firstName = $partesNome[0]
        $lastName = $partesNome[$partesNome.Length - 1]
        $userPrincipalName = "$sam$($comboBoxUPNSuffix.SelectedItem)"

        # Validação básica
        if ([string]::IsNullOrWhiteSpace($nome) -or [string]::IsNullOrWhiteSpace($sam) -or [string]::IsNullOrWhiteSpace($textBoxes["Senha"].Text)) {
            [System.Windows.Forms.MessageBox]::Show("Preencha todos os campos obrigatórios!", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        try {
            New-ADUser `
                -Name $nome `
                -GivenName $firstName `
                -Surname $lastName `
                -SamAccountName $sam `
                -UserPrincipalName $userPrincipalName `
                -AccountPassword $senha `
                -DisplayName $nome `
                -EmailAddress $email `
                -Department $departamento `
                -Title $cargo `
                -Path $pathOU `
                -Enabled $true `
                -ChangePasswordAtLogon $true

            if ($grupo -ne "") {
                $grupos = $grupo -split "[,;\n]"
                foreach ($g in $grupos) {
                    $gTrimmed = $g.Trim()
                    if ($gTrimmed -ne "") {
                        Add-ADGroupMember -Identity $gTrimmed -Members $sam
                    }
                }
            }

            [System.Windows.Forms.MessageBox]::Show("Usuário '$nome' criado com sucesso.","Sucesso",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
            $form.Close()
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Erro ao criar usuário: $_","Erro",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $form.Controls.Add($button)

    $form.ShowDialog()
}

function Abrir-GerenciarUsuario {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName PresentationFramework
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue

    # --- Interface Gráfica ---
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Gerenciar Usuário AD"
    $form.Size = New-Object System.Drawing.Size(350,320)
    $form.StartPosition = "CenterScreen"

    # Rótulo
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Digite o nome de login do usuário:"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(10,20)
    $form.Controls.Add($label)

    # Caixa de texto para nome de usuário
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Size = New-Object System.Drawing.Size(300,20)
    $textBox.Location = New-Object System.Drawing.Point(10,45)
    $form.Controls.Add($textBox)

    # Label para mostrar o status
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Status: Não verificado"
    $statusLabel.AutoSize = $true
    $statusLabel.Location = New-Object System.Drawing.Point(10,75)
    $form.Controls.Add($statusLabel)

    # Botão: Verificar Status
    $btnVerificar = New-Object System.Windows.Forms.Button
    $btnVerificar.Text = "Verificar Status"
    $btnVerificar.Size = New-Object System.Drawing.Size(300,30)
    $btnVerificar.Location = New-Object System.Drawing.Point(10,100)
    $btnVerificar.Add_Click({
        $usuario = $textBox.Text
        if ([string]::IsNullOrWhiteSpace($usuario)) {
            [System.Windows.Forms.MessageBox]::Show("Digite um nome de usuário!", "Erro")
            return
        }
        
        try {
            $user = Get-ADUser -Identity $usuario -Properties Enabled
            if ($user.Enabled) {
                $statusLabel.Text = "Status: ATIVO (Habilitado)"
                $statusLabel.ForeColor = [System.Drawing.Color]::Green
            } else {
                $statusLabel.Text = "Status: INATIVO (Desabilitado)"
                $statusLabel.ForeColor = [System.Drawing.Color]::Red
            }
            [System.Windows.Forms.MessageBox]::Show("Status do usuário verificado com sucesso!", "Sucesso")
        } catch {
            $statusLabel.Text = "Status: Usuário não encontrado"
            $statusLabel.ForeColor = [System.Drawing.Color]::DarkRed
            [System.Windows.Forms.MessageBox]::Show("Erro ao verificar usuário: $_", "Erro")
        }
    })
    $form.Controls.Add($btnVerificar)

    # Botão: Desabilitar usuário
    $btnDesabilitar = New-Object System.Windows.Forms.Button
    $btnDesabilitar.Text = "Desabilitar"
    $btnDesabilitar.Size = New-Object System.Drawing.Size(145,40)
    $btnDesabilitar.Location = New-Object System.Drawing.Point(165,140)
    $btnDesabilitar.BackColor = [System.Drawing.Color]::LightCoral
    $btnDesabilitar.Add_Click({
        $usuario = $textBox.Text
        if ([string]::IsNullOrWhiteSpace($usuario)) {
            [System.Windows.Forms.MessageBox]::Show("Digite um nome de usuário!", "Erro")
            return
        }
        
        try {
            Disable-ADAccount -Identity $usuario
            $statusLabel.Text = "Status: INATIVO (Desabilitado)"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
            [System.Windows.Forms.MessageBox]::Show("Usuário desabilitado com sucesso!", "Sucesso")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erro ao desabilitar usuário: $_", "Erro")
        }
    })
    $form.Controls.Add($btnDesabilitar)

    # Botão: Habilitar usuário
    $btnHabilitar = New-Object System.Windows.Forms.Button
    $btnHabilitar.Text = "Habilitar"
    $btnHabilitar.Size = New-Object System.Drawing.Size(145,40)
    $btnHabilitar.Location = New-Object System.Drawing.Point(10,140)
    $btnHabilitar.BackColor = [System.Drawing.Color]::LightGreen
    $btnHabilitar.Add_Click({
        $usuario = $textBox.Text
        if ([string]::IsNullOrWhiteSpace($usuario)) {
            [System.Windows.Forms.MessageBox]::Show("Digite um nome de usuário!", "Erro")
            return
        }
        
        try {
            Enable-ADAccount -Identity $usuario
            $statusLabel.Text = "Status: ATIVO (Habilitado)"
            $statusLabel.ForeColor = [System.Drawing.Color]::Green
            [System.Windows.Forms.MessageBox]::Show("Usuário habilitado com sucesso!", "Sucesso")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erro ao habilitar usuário: $_", "Erro")
        }
    })
    $form.Controls.Add($btnHabilitar)

    # Botão: Excluir usuário
    $btnExcluir = New-Object System.Windows.Forms.Button
    $btnExcluir.Text = "Excluir Usuário"
    $btnExcluir.Size = New-Object System.Drawing.Size(300,40)
    $btnExcluir.Location = New-Object System.Drawing.Point(10,190)
    $btnExcluir.Add_Click({
        $usuario = $textBox.Text
        if ([string]::IsNullOrWhiteSpace($usuario)) {
            [System.Windows.Forms.MessageBox]::Show("Digite um nome de usuário!", "Erro")
            return
        }
        $confirmar = [System.Windows.Forms.MessageBox]::Show("Tem certeza que deseja excluir PERMANENTEMENTE o usuário $usuario?", "Confirmar Exclusão", "YesNo", [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirmar -eq "Yes") {
            try {
                Remove-ADUser -Identity $usuario -Confirm:$false
                [System.Windows.Forms.MessageBox]::Show("Usuário $usuario excluído com sucesso!", "Sucesso")
                $form.Close()
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Erro ao excluir: $_", "Erro")
            }
        }
    })
    $form.Controls.Add($btnExcluir)

    # Rodar o formulário
    [void]$form.ShowDialog()
}

function Abrir-AlterarSenha {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName PresentationFramework

    # --- Interface Gráfica ---
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Alterar Senha do Usuário AD"
    $form.Size = New-Object System.Drawing.Size(300,250)
    $form.StartPosition = "CenterScreen"

    # Rótulo para nome de usuário
    $labelUsuario = New-Object System.Windows.Forms.Label
    $labelUsuario.Text = "Nome de usuário (login):"
    $labelUsuario.AutoSize = $true
    $labelUsuario.Location = New-Object System.Drawing.Point(10,20)
    $form.Controls.Add($labelUsuario)

    # Caixa de texto para nome de usuário
    $textBoxUsuario = New-Object System.Windows.Forms.TextBox
    $textBoxUsuario.Size = New-Object System.Drawing.Size(250,20)
    $textBoxUsuario.Location = New-Object System.Drawing.Point(10,45)
    $form.Controls.Add($textBoxUsuario)

    # Rótulo para nova senha
    $labelSenha = New-Object System.Windows.Forms.Label
    $labelSenha.Text = "Nova senha:"
    $labelSenha.AutoSize = $true
    $labelSenha.Location = New-Object System.Drawing.Point(10,75)
    $form.Controls.Add($labelSenha)

    # Caixa de texto para nova senha
    $textBoxSenha = New-Object System.Windows.Forms.TextBox
    $textBoxSenha.Size = New-Object System.Drawing.Size(250,20)
    $textBoxSenha.Location = New-Object System.Drawing.Point(10,100)
    $textBoxSenha.UseSystemPasswordChar = $true
    $form.Controls.Add($textBoxSenha)

    # Botão: Alterar senha
    $btnAlterar = New-Object System.Windows.Forms.Button
    $btnAlterar.Text = "Alterar Senha"
    $btnAlterar.Size = New-Object System.Drawing.Size(250,40)
    $btnAlterar.Location = New-Object System.Drawing.Point(10,140)
    $btnAlterar.Add_Click({
        $usuario = $textBoxUsuario.Text
        $novaSenha = $textBoxSenha.Text
        
        if ([string]::IsNullOrWhiteSpace($usuario) -or [string]::IsNullOrWhiteSpace($novaSenha)) {
            [System.Windows.Forms.MessageBox]::Show("Preencha todos os campos!", "Erro")
            return
        }
        
        try {
            $securePassword = ConvertTo-SecureString $novaSenha -AsPlainText -Force
            Set-ADAccountPassword -Identity $usuario -NewPassword $securePassword -Reset
            [System.Windows.Forms.MessageBox]::Show("Senha alterada com sucesso para o usuário $usuario!", "Sucesso")
            $form.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erro ao alterar senha: $_", "Erro")
        }
    })
    $form.Controls.Add($btnAlterar)

    # Rodar o formulário
    [void]$form.ShowDialog()
}

function Instalar-RSAT {
    # Verifica se o módulo ActiveDirectory já está disponível
    if (Get-Module -Name ActiveDirectory -ListAvailable) {
        [System.Windows.Forms.MessageBox]::Show("O módulo Active Directory (RSAT) já está instalado.", "RSAT Instalado", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # Verifica se é Windows 10/11 ou Windows Server
    $isClientOS = (Get-CimInstance -ClassName Win32_OperatingSystem).ProductType -eq 1

    if ($isClientOS) {
        # Para Windows 10/11
        try {
            # Verifica se o recurso RSAT está disponível
            $feature = Get-WindowsCapability -Name Rsat.ActiveDirectory* -Online -ErrorAction Stop
            
            if ($feature.State -eq "Installed") {
                [System.Windows.Forms.MessageBox]::Show("O RSAT para Active Directory já está instalado.", "RSAT Instalado", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                return
            }
            
            # Tenta instalar o RSAT
            $result = Add-WindowsCapability -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -Online -ErrorAction Stop
            
            if ($result.RestartNeeded) {
                [System.Windows.Forms.MessageBox]::Show("RSAT instalado com sucesso! Uma reinicialização é necessária.", "Sucesso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                [System.Windows.Forms.MessageBox]::Show("RSAT instalado com sucesso!", "Sucesso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erro ao instalar RSAT: $_`n`nPor favor, instale manualmente através do 'Gerenciar recursos opcionais' no Windows.", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        # Para Windows Server
        try {
            if ((Get-WindowsFeature -Name RSAT-AD-PowerShell).Installed) {
                [System.Windows.Forms.MessageBox]::Show("O RSAT para Active Directory já está instalado.", "RSAT Instalado", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                return
            }
            
            # Instala o módulo AD no Server
            Install-WindowsFeature -Name RSAT-AD-PowerShell -IncludeManagementTools -ErrorAction Stop
            [System.Windows.Forms.MessageBox]::Show("RSAT instalado com sucesso!", "Sucesso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erro ao instalar RSAT: $_", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# --- Interface Principal ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Menu Principal - Gerenciamento AD"
$form.Size = New-Object System.Drawing.Size(360,350)  # Aumentado para acomodar o novo botão
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Text = "______________________________________________________"
$label.Location = New-Object System.Drawing.Point(10,95)
$label.AutoSize = $true
$form.Controls.Add($label)

$btnRSAT = New-Object System.Windows.Forms.Button
$btnRSAT.Text = "Verificar / Instalar RSAT"
$btnRSAT.Size = New-Object System.Drawing.Size(300,40)
$btnRSAT.Location = New-Object System.Drawing.Point(20,40)
$btnRSAT.Add_Click({ Instalar-RSAT })
$form.Controls.Add($btnRSAT)

$btnCriar = New-Object System.Windows.Forms.Button
$btnCriar.Text = "Criar Usuário AD"
$btnCriar.Size = New-Object System.Drawing.Size(300,40)
$btnCriar.Location = New-Object System.Drawing.Point(20,140)
$btnCriar.Add_Click({ Abrir-CriacaoUsuario })
$form.Controls.Add($btnCriar)

$btnGerenciar = New-Object System.Windows.Forms.Button
$btnGerenciar.Text = "Desativar/Excluir Usuário"
$btnGerenciar.Size = New-Object System.Drawing.Size(300,40)
$btnGerenciar.Location = New-Object System.Drawing.Point(20,190)
$btnGerenciar.Add_Click({ Abrir-GerenciarUsuario })
$form.Controls.Add($btnGerenciar)

$btnAlterarSenha = New-Object System.Windows.Forms.Button
$btnAlterarSenha.Text = "Alterar Senha"
$btnAlterarSenha.Size = New-Object System.Drawing.Size(300,40)
$btnAlterarSenha.Location = New-Object System.Drawing.Point(20,240)
$btnAlterarSenha.Add_Click({ Abrir-AlterarSenha })
$form.Controls.Add($btnAlterarSenha)

$form.ShowDialog()
