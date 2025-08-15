# Define the SMTP server and sender email
$smtpServer = "smtp.gmail.com"  # Gmail's SMTP server
$senderEmail = "deandregraciano1995@gmail.com"  # Replace with your Gmail address
$port = 587  # Port for TLS encryption

# Prompt for Gmail credentials (use your email and App Password if 2FA is enabled)
$credential = Get-Credential  # This will prompt you for your Gmail email and password (or App Password)

# Define the list of incidents with details and personalized email bodies
$incidents = @(
    @{
        Incident = "123456"
        Name = "Dhayana"
        Reason = "The document was voided due to missing or incomplete information."
        Email = "dhayana.reinoso@gmail.com"
        Body = @"
Hi Dhayana,

I hope this message finds you well. I just wanted to let you know that the document associated with Incident #123456 has been voided.

**Reason for Void:**  
The document was voided due to missing or incomplete information.

If you have any questions or need help sorting this out, feel free to reach out.

Have a great day!  

Best regards,  
DeAndre
"@
    },
    @{
        Incident = "654321"
        Name = "DeAndre"
        Reason = "The document was voided due to missing signatures."
        Email = "deandre.graciano1995@gmail.com"
        Body = @"
Hi DeAndre,

I hope this message finds you well. I just wanted to let you know that the document associated with Incident #654321 has been voided.

**Reason for Void:**  
The document was voided due to missing signatures.

If you have any questions or need help sorting this out, feel free to reach out.

Have a great day!  

Best regards,  
Your Name
"@
    }
)

# Loop through each incident and send an email
foreach ($incident in $incidents) {
    $to = $incident.Email
    $subject = "Follow-Up: Document Incident #$($incident.Incident) Voided"
    $body = $incident.Body

    try {
        # Send the email
        Send-MailMessage -From $senderEmail `
                         -To $to `
                         -Subject $subject `
                         -Body $body `
                         -SmtpServer $smtpServer `
                         -Port $port `
                         -Credential $credential `
                         -UseSsl `
                         -BodyAsHtml
        Write-Host "Email successfully sent to $to for Incident #$($incident.Incident)"
    } catch {
        # Handle errors
        Write-Host "Failed to send email to $to for Incident #$($incident.Incident): $_"
    }
}