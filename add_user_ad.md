
```
# Установка RSAT: Active Directory Tools
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.Tools~~~~0.0.1.0 -ErrorAction Stop

# Проверка и установка политики выполнения
$currentExecutionPolicy = Get-ExecutionPolicy
if ($currentExecutionPolicy -ne "RemoteSigned" -and $currentExecutionPolicy -ne "Unrestricted") {
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Reflection.Assembly]::LoadWithPartialName('System.Web') | Out-Null

# Функция для проверки наличия модуля ActiveDirectory
function Check-ActiveDirectoryModule {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# Функция для установки модуля ActiveDirectory
function Install-ActiveDirectoryModule {
    try {
        Write-Host "Модуль ActiveDirectory не найден. Попытка установить..."
        Install-WindowsFeature RSAT-AD-PowerShell -ErrorAction Stop
        return $true
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Не удалось установить модуль ActiveDirectory. Убедитесь, что у вас есть права администратора.", "Ошибка установки", "OK", "Error")
        return $false
    }
}

# Проверяем, установлен ли модуль ActiveDirectory
if (-not (Check-ActiveDirectoryModule)) {
    if (-not (Install-ActiveDirectoryModule)) {
        exit
    }
}

# Функция для проверки подключения к AD
function Test-ADConnection {
    param(
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        $null = Get-ADDomain -Credential $Credential -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# Функция для запроса учетных данных с проверкой
function Get-ADCredentials {
    do {
        $credForm = New-Object Windows.Forms.Form
        $credForm.Text = "Подключение к Active Directory"
        $credForm.Size = New-Object Drawing.Size(350, 200)
        $credForm.StartPosition = "CenterScreen"
        $credForm.FormBorderStyle = "FixedDialog"
        $credForm.MaximizeBox = $false

        $labelDomain = New-Object Windows.Forms.Label
        $labelDomain.Text = "Домен:"
        $labelDomain.Location = New-Object Drawing.Point(20, 20)
        $labelDomain.Size = New-Object Drawing.Size(100, 20)

        $textBoxDomain = New-Object Windows.Forms.TextBox
        $textBoxDomain.Location = New-Object Drawing.Point(120, 20)
        $textBoxDomain.Size = New-Object Drawing.Size(180, 20)
        $textBoxDomain.Text = "YOUR_DOMAIN.local"

        $labelUsername = New-Object Windows.Forms.Label
        $labelUsername.Text = "Пользователь:"
        $labelUsername.Location = New-Object Drawing.Point(20, 50)
        $labelUsername.Size = New-Object Drawing.Size(100, 20)

        $textBoxUsername = New-Object Windows.Forms.TextBox
        $textBoxUsername.Location = New-Object Drawing.Point(120, 50)
        $textBoxUsername.Size = New-Object Drawing.Size(180, 20)
        $textBoxUsername.Text = "$env:USERNAME@YOUR_DOMAIN.local"

        $labelPassword = New-Object Windows.Forms.Label
        $labelPassword.Text = "Пароль:"
        $labelPassword.Location = New-Object Drawing.Point(20, 80)
        $labelPassword.Size = New-Object Drawing.Size(100, 20)

        $textBoxPassword = New-Object Windows.Forms.TextBox
        $textBoxPassword.Location = New-Object Drawing.Point(120, 80)
        $textBoxPassword.Size = New-Object Drawing.Size(180, 20)
        $textBoxPassword.PasswordChar = '*'

        $buttonOK = New-Object Windows.Forms.Button
        $buttonOK.Text = "OK"
        $buttonOK.Location = New-Object Drawing.Point(120, 120)
        $buttonOK.Size = New-Object Drawing.Size(80, 30)
        $buttonOK.DialogResult = [Windows.Forms.DialogResult]::OK

        $buttonCancel = New-Object Windows.Forms.Button
        $buttonCancel.Text = "Отмена"
        $buttonCancel.Location = New-Object Drawing.Point(220, 120)
        $buttonCancel.Size = New-Object Drawing.Size(80, 30)
        $buttonCancel.DialogResult = [Windows.Forms.DialogResult]::Cancel

        $credForm.Controls.AddRange(@(
            $labelDomain, $textBoxDomain,
            $labelUsername, $textBoxUsername,
            $labelPassword, $textBoxPassword,
            $buttonOK, $buttonCancel
        ))

        $credForm.AcceptButton = $buttonOK
        $credForm.CancelButton = $buttonCancel

        $result = $credForm.ShowDialog()

        if ($result -eq [Windows.Forms.DialogResult]::Cancel) {
            return $null
        }

        $username = $textBoxUsername.Text
        if (-not $username.Contains("@")) {
            $username = "$username@$($textBoxDomain.Text)"
        }

        $securePassword = ConvertTo-SecureString $textBoxPassword.Text -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($username, $securePassword)
        
        # Проверка подключения к AD
        if (Test-ADConnection -Credential $credential) {
            return $credential
        } else {
            $retryResult = [Windows.Forms.MessageBox]::Show(
                "Неверное имя пользователя или пароль. Повторить ввод?", 
                "Ошибка аутентификации", 
                "RetryCancel", 
                "Error"
            )
            
            if ($retryResult -eq "Cancel") {
                return $null
            }
        }
    } while ($true)
}

# Получаем учетные данные
$ADCredential = Get-ADCredentials
if (-not $ADCredential) {
    exit
}

# Остальная часть скрипта для создания нового пользователя
$form = New-Object Windows.Forms.Form
$form.Text = "Создание пользователя в Active Directory"
$form.Size = New-Object Drawing.Size(450, 675) # Размер окна x,y
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$labelHeader = New-Object Windows.Forms.Label
$labelHeader.Text = "Введите данные нового пользователя"
$labelHeader.Font = New-Object Drawing.Font("Arial", 12, [Drawing.FontStyle]::Bold)
$labelHeader.Location = New-Object Drawing.Point(20, 20)
$labelHeader.Size = New-Object Drawing.Size(400, 30)

$form.Controls.Add($labelHeader)

$labels = @("Логин", "Имя", "Фамилия", "Отчество", "Должность", "Отдел", "Компания", "Подразделение (OU)")
$textBoxes = @()

for ($i = 0; $i -lt $labels.Count; $i++) {
    $yPos = 60 + $i * 50
    
    $label = New-Object Windows.Forms.Label
    $label.Text = $labels[$i]
    $label.Location = New-Object Drawing.Point(20, $yPos)
    $label.Size = New-Object Drawing.Size(150, 20)
    
    $textBox = New-Object Windows.Forms.TextBox
    $textBox.Location = New-Object Drawing.Point(180, $yPos)
    $textBox.Size = New-Object Drawing.Size(230, 20)
    
    if ($labels[$i] -eq "Подразделение (OU)") {
        $textBox.Text = "OU=YOUR_OU,OU=YOUR_OU,OU=YOUR_OU,DC=YOUR_DOMAIN,DC=local"
    }
    
    $form.Controls.Add($label)
    $form.Controls.Add($textBox)
    $textBoxes += $textBox
}

$checkBox = New-Object Windows.Forms.CheckBox
$checkBox.Text = "Требовать смену пароля при первом входе"
$checkBox.Location = New-Object Drawing.Point(20, 460)
$checkBox.Size = New-Object Drawing.Size(300, 20)
$checkBox.Checked = $true

$labelPassword = New-Object Windows.Forms.Label
$labelPassword.Text = "Пароль:"
$labelPassword.Location = New-Object Drawing.Point(20, 490)

$textBoxPassword = New-Object Windows.Forms.TextBox
$textBoxPassword.Location = New-Object Drawing.Point(180, 490)
$textBoxPassword.Size = New-Object Drawing.Size(230, 20)
$textBoxPassword.PasswordChar = '*'

$buttonGenerate = New-Object Windows.Forms.Button
$buttonGenerate.Text = "Сгенерировать"
$buttonGenerate.Location = New-Object Drawing.Point(180, 520)
$buttonGenerate.Size = New-Object Drawing.Size(100, 30)
$buttonGenerate.Add_Click({
    $global:GeneratedPassword = [System.Web.Security.Membership]::GeneratePassword(12, 3)
    $textBoxPassword.Text = $global:GeneratedPassword
})

$buttonCreate = New-Object Windows.Forms.Button
$buttonCreate.Text = "Создать пользователя"
$buttonCreate.Location = New-Object Drawing.Point(150, 570)
$buttonCreate.Size = New-Object Drawing.Size(150, 30)
$buttonCreate.Add_Click({
    try {
        $Username = $textBoxes[0].Text
        $FirstName = $textBoxes[1].Text
        $LastName = $textBoxes[2].Text
        $MiddleName = $textBoxes[3].Text
        $Title = $textBoxes[4].Text
        $Department = $textBoxes[5].Text
        $Company = $textBoxes[6].Text
        $OU = $textBoxes[7].Text
        
        $PasswordText = if ($textBoxPassword.Text) { $textBoxPassword.Text } else { $global:GeneratedPassword }
        $Password = ConvertTo-SecureString $PasswordText -AsPlainText -Force
        
        $ChangePasswordAtLogon = $checkBox.Checked
        
        $Initials = "$($FirstName[0])$($MiddleName[0])"
        $DisplayName = "$LastName $FirstName $MiddleName"

        if (Get-ADUser -Filter "SamAccountName -eq '$Username'" -Credential $ADCredential -ErrorAction SilentlyContinue) {
            [Windows.Forms.MessageBox]::Show("Пользователь $Username уже существует!", "Ошибка", "OK", "Error")
            return
        }

        if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$OU'" -Credential $ADCredential -ErrorAction SilentlyContinue)) {
            [Windows.Forms.MessageBox]::Show("Подразделение $OU не найдено!", "Ошибка", "OK", "Error")
            return
        }

        $NewUserParams = @{
            Name                  = $DisplayName
            SamAccountName       = $Username
            UserPrincipalName    = "$Username@YOUR_DOMAIN.local"
            GivenName           = $FirstName
            Surname             = $LastName
            DisplayName         = $DisplayName
            Initials            = $Initials
            Title               = $Title
            Department          = $Department
            Company             = $Company
            AccountPassword     = $Password
            Enabled             = $true
            Path                = $OU
            ChangePasswordAtLogon = $ChangePasswordAtLogon
            EmailAddress        = "$Username@YOUR_DOMAIN.local"
            Credential          = $ADCredential
            ErrorAction         = 'Stop'
        }

        $NewUser = New-ADUser @NewUserParams

        $result = [Windows.Forms.MessageBox]::Show(
            "Пользователь $Username успешно создан!`n`nЛогин: $Username`nПароль: $PasswordText`n`nСкопировать пароль в буфер обмена?", 
            "Успех", 
            "YesNo", 
            "Information"
        )
        
        if ($result -eq "Yes") {
            [Windows.Forms.Clipboard]::SetText($PasswordText)
        }
        
    } catch {
        [Windows.Forms.MessageBox]::Show("Ошибка при создании пользователя: $_", "Ошибка", "OK", "Error")
    }
})

$form.Controls.Add($checkBox)
$form.Controls.Add($labelPassword)
$form.Controls.Add($textBoxPassword)
$form.Controls.Add($buttonGenerate)
$form.Controls.Add($buttonCreate)

# Добавляем надпись внизу формы
$labelFooter = New-Object Windows.Forms.Label
$labelFooter.Text = "Сделано @MikeWazowskai"
$labelFooter.Font = New-Object Drawing.Font("Arial", 8, [Drawing.FontStyle]::Italic)
$labelFooter.Location = New-Object Drawing.Point(295, 610)
$labelFooter.Size = New-Object Drawing.Size(250, 20)

$form.Controls.Add($labelFooter)

$form.ShowDialog()
```