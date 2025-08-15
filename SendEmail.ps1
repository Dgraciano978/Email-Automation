# Define the SMTP server and sender email
$smtpServer = "smtp.yourdomain.com"  # Replace with your SMTP server
$senderEmail = "your-email@yourdomain.com"  # Replace with your email address

# Define the list of incidents with details and personalized email bodies
$incidents = @(
    @{
        Incident = "458169"
        Name = "Phay"
        Reason = "The document was marked as 'VOID' due to incomplete or missing vendor onboarding information, such as banking details or other required forms."
        Email = "phay@mimecast.com"
        Body = @"
Hi Phay,

I hope this message finds you well. I just wanted to let you know that the document associated with Incident #458169 has been voided.

**Reason for Void:**  
The document was marked as 'VOID' due to incomplete or missing vendor onboarding information, such as banking details or other required forms.

If you have any questions or need help sorting this out, feel free to reach out.

Have a great day!  

Best regards,  
Deandre
"@
    },
    @{
        Incident = "455486"
        Name = "Redwards"
        Reason = "The document was voided because the required tax residency certificate, MSME certificate, PAN card, or other regulatory documents were not uploaded or completed."
        Email = "redwards@mimecast.com"
        Body = @"
Hi Redwards,

I hope this message finds you well. I wanted to let you know that the document related to Incident #455486 has been voided.

**Reason for Void:**  
The document couldn’t be processed because the required tax residency certificate, MSME certificate, PAN card, or other regulatory documents weren’t uploaded or completed.

If you need assistance or have any questions, don’t hesitate to reach out.

Have a great day!  

Best regards,  
Deandre
"@
    },
    @{
        Incident = "458734"
        Name = "Nmunsami"
        Reason = "The W-9 form was voided due to incomplete or incorrect taxpayer information or failure to meet compliance requirements."
        Email = "nmunsami@mimecast.com"
        Body = @"
Hi Nmunsami,

I hope this message finds you well. I wanted to quickly let you know that the document associated with Incident #458734 has been voided.

**Reason for Void:**  
It seems the W-9 form wasn’t processed because of incomplete or incorrect taxpayer information or compliance issues.

Let me know if you need help addressing this or have any questions.

Have a great day!  

Best regards,  
Deandre
"@
    },
    @{
        Incident = "458443"
        Name = "Efisher"
        Reason = "The W-9 form was voided due to incomplete or incorrect taxpayer information or failure to meet compliance requirements."
        Email = "efisher@mimecast.com"
        Body = @"
Hi Efisher,

I hope this message finds you well. Just a quick update regarding the document tied to Incident #458443 — it’s been voided.

**Reason for Void:**  
The W-9 form had some incomplete or incorrect taxpayer information, which caused some compliance issues.

If there’s anything I can do to assist, just let me know.

Have a great day!  

Best regards,  
Deandre
"@
    },
    @{
        Incident = "459942"
        Name = "Aokwudinka"
        Reason = "The document was voided due to incomplete vendor onboarding information, such as missing banking details or incomplete W-9 forms."
        Email = "aokwudinka@mimecast.com"
        Body = @"
Hi Aokwudinka,

I hope this message finds you well. I wanted to let you know that the document for Incident #459942 has been voided.

**Reason for Void:**  
It looks like the vendor onboarding information wasn’t fully completed — there might be missing banking details or an incomplete W-9 form.

Feel free to reach out if you have any questions or need help fixing this.

Have a great day!  

Best regards,  
Deandre
"@
    },
    @{
        Incident = "459943"
        Name = "Aokwudinka"
        Reason = "The document was voided as no signatures were completed. The envelope status indicates it was voided by the originator."
        Email = "aokwudinka@mimecast.com"
        Body = @"
Hi Aokwudinka,

I hope this message finds you well. I just wanted to follow up about Incident #459943. The document has been voided because no signatures were completed.

If you need help resolving this or have any questions, just let me know.

Have a great day!  

Best regards,  
Deandre
"@
    },
    @{
        Incident = "460149"
        Name = "Lreddy"
        Reason = "The document was voided due to incomplete submission of required tax documentation (e.g., Tax Residency Certificate, PAN Card, etc.) for compliance with Indian tax regulations."
        Email = "lreddy@mimecast.com"
        Body = @"
Hi Lreddy,

I hope this message finds you well. I’m reaching out to let you know that the document for Incident #460149 has been voided.

**Reason for Void:**  
It seems the submission of required tax documentation (e.g., Tax Residency Certificate, PAN Card, etc.) wasn’t completed, which caused compliance issues.

If there’s anything I can do to assist, just let me know!

Have a great day!  

Best regards,  
Deandre
"@
    },
    @{
        Incident = "461053"
        Name = "Dbest"
        Reason = "The document was voided because the vendor onboarding process was incomplete, specifically the submission of banking details and other required forms."
        Email = "dbest@mimecast.com"
        Body = @"
Hi Dbest,

I hope this message finds you well. I wanted to let you know that the document tied to Incident #461053 has been voided.

**Reason for Void:**  
The vendor onboarding process wasn’t fully completed — banking details or other required forms might be missing.

If you have any questions or need help resolving this, just let me know.

Have a great day!  

Best regards,  
Deandre
"@
    },
    @{
        Incident = "462213"
        Name = "Mmasciarelli"
        Reason = "The document was voided due to incomplete vendor onboarding, including missing banking information and tax forms."
        Email = "mmasciarelli@mimecast.com"
        Body = @"
Hi Mmasciarelli,

I hope this message finds you well. I wanted to follow up and let you know that the document for Incident #462213 has been voided.

**Reason for Void:**  
It seems the vendor onboarding process is incomplete, with missing banking information or tax forms.

If there’s anything I can help with, please don’t hesitate to reach out.

Have a great day!  

Best regards,  
Deandre
"@
    },
    @{
        Incident = "464940"
        Name = "Jil Augel"
        Reason = "The document was voided because it was incomplete or not signed. The certificate of completion indicates no signatures were captured."
        Email = "jaugel@mimecast.com"
        Body = @"
Hi Jil,

I hope this message finds you well. I wanted to let you know that the document related to Incident #464940 has been voided.

**Reason for Void:**  
The document couldn’t be completed because it was either incomplete or didn’t have the required signatures.

Let me know if you have any questions or need help with this.

Have a great day!  

Best regards,  
Deandre
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