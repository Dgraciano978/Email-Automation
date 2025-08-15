# Define the SMTP server and sender email
$smtpServer = "smtp.example.com"  # Replace with your SMTP server
$senderEmail = "your-email@example.com"  # Replace with your email address

# Define the list of incidents with details and personalized email bodies
$incidents = @(
    @{
        Incident = "123456"
        Name = "John"
        Reason = "The document was voided due to missing or incomplete information."
        Email = "john.doe@example.com"
        Body = @"
Hi John,

I hope this message finds you well. I just wanted to let you know that the document associated with Incident #123456 has been voided.

**Reason for Void:**  
The document was voided due to missing or incomplete information.

If you have any questions or need help sorting this out, feel free to reach out.

Have a great day!  

Best regards,  
Your Name
"@
    },
    @{
        Incident = "654321"
        Name = "Jane"
        Reason = "The document was voided due to missing signatures."
        Email = "jane.doe@example.com"
        Body = @"
Hi Jane,

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

    # Send the email
    Send-MailMessage -From $senderEmail -To $to -Subject $subject -Body $body -SmtpServer $smtpServer -BodyAsHtml
    Write-Host "Email sent to $to for Incident #$($incident.Incident)"
}