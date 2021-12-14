<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    MailClient
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '600,300'
$Form.text                       = "Send Mail Client"
$Form.TopMost                    = $false
$Form.MaximizeBox                = $false
$Form.FormBorderStyle            = 'Fixed3D'

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Mail Server"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(30,41)
$Label1.Font                     = 'Microsoft Sans Serif,10'

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "From"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(34,77)
$Label2.Font                     = 'Microsoft Sans Serif,10'

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "To"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(36,110)
$Label3.Font                     = 'Microsoft Sans Serif,10'

$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = "Anzahl"
$Label4.AutoSize                 = $true
$Label4.width                    = 25
$Label4.height                   = 10
$Label4.location                 = New-Object System.Drawing.Point(33,151)
$Label4.Font                     = 'Microsoft Sans Serif,10'

$Label5                          = New-Object system.Windows.Forms.Label
$Label5.text                     = "Prio High"
$Label5.AutoSize                 = $true
$Label5.width                    = 25
$Label5.height                   = 10
$Label5.location                 = New-Object System.Drawing.Point(32,192)
$Label5.Font                     = 'Microsoft Sans Serif,10'

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 135
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(106,36)
$TextBox1.Font                   = 'Microsoft Sans Serif,10'

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $false
$TextBox2.width                  = 135
$TextBox2.height                 = 20
$TextBox2.location               = New-Object System.Drawing.Point(105,75)
$TextBox2.Font                   = 'Microsoft Sans Serif,10'

$TextBox3                        = New-Object system.Windows.Forms.TextBox
$TextBox3.multiline              = $false
$TextBox3.width                  = 135
$TextBox3.height                 = 20
$TextBox3.location               = New-Object System.Drawing.Point(105,109)
$TextBox3.Font                   = 'Microsoft Sans Serif,10'

$TextBox4                        = New-Object system.Windows.Forms.TextBox
$TextBox4.multiline              = $false
$TextBox4.width                  = 30
$TextBox4.height                 = 20
$TextBox4.location               = New-Object System.Drawing.Point(106,152)
$TextBox4.Font                   = 'Microsoft Sans Serif,10'

$CheckBox1                       = New-Object system.Windows.Forms.CheckBox
$CheckBox1.AutoSize              = $false
$CheckBox1.width                 = 95
$CheckBox1.height                = 20
$CheckBox1.location              = New-Object System.Drawing.Point(105,190)
$CheckBox1.Font                  = 'Microsoft Sans Serif,10'

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "Send Mail"
$Button1.width                   = 220
$Button1.height                  = 57
$Button1.location                = New-Object System.Drawing.Point(205,221)
$Button1.Font                    = 'Microsoft Sans Serif,10'

$TextBox5                        = New-Object system.Windows.Forms.TextBox
$TextBox5.multiline              = $true
$TextBox5.width                  = 299
$TextBox5.height                 = 162
$TextBox5.enabled                = $false
$TextBox5.location               = New-Object System.Drawing.Point(270,32)
$TextBox5.Font                   = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($Label1,$Label2,$Label3,$Label4,$Label5,$TextBox1,$TextBox2,$TextBox3,$TextBox4,$CheckBox1,$Button1,$TextBox5))

$Button1.Add_Click({ sendMail })

$TextBox1.Text = "172.30.1.1" 
$TextBox2.Text = "hello@test.ch"
$TextBox3.Text = ""
$TextBox4.Text = "1"


function sendMail {
   $mailserver=""
   $from=""
   $to=""
   $anzahl=""
   $prio=""
    
   $mailserver = $TextBox1.Text
   $from       = $TextBox2.Text
   $to         = $TextBox3.Text
   $anzahl     = $TextBox4.Text

   $prio       = $CheckBox1.Checked
   if ($prio){
        $prio = "High"
   } else {
        $prio = "Low"
   }
     
         
    For ($i=1; $i -le $anzahl; $i++) {
        $MailMessage = @{

        from        = $from
        to          = $to
        subject     = "$i - If you're happy and you know it"
        body      = @'
If you're happy and you know it, clap your hands.
If you're happy and you know it, clap your hands.
If you're happy and you know it
and you really want to show it,
If you're happy and you know it, clap your hands.

If you're happy and you know it, slap your sides.
If you're happy and you know it, slap your sides.
If you're happy and you know it
and you really want to show it,
If you're happy and you know it, slap your sides.

If you're happy and you know it, stomp your feet.
If you're happy and you know it, stomp your feet.
If you're happy and you know it
and you really want to show it,
If you're happy and you know it, stomp your feet.

If you're happy and you know it, snap your fingers.
If you're happy and you know it, snap your fingers.
If you're happy and you know it
and you really want to show it,
If you're happy and you know it, snap your fingers.

If you're happy and you know it, sniff your nose.
If you're happy and you know it, sniff your nose.
If you're happy and you know it
and you really want to show it,
If you're happy and you know it, sniff your nose.

If you're happy and you know it, shout "We are".
If you're happy and you know it, shout "We are".
If you're happy and you know it
and you really want to show it,
If you're happy and you know it, shout "We are".
'@
        smtpserver  = $mailserver
        Priority = [System.Net.Mail.MailPriority]::$prio
        }
        
        Invoke-Expression "Send-MailMessage @MailMessage" -ErrorVariable errorMessage
        if ($errorMessage) {
        $TextBox5.Text = $errorMessage
        
        } else {
        $TextBox5.Text = "Mail sent to $to"
        
        }

        }
}


[void]$Form.ShowDialog()