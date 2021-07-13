$path = "C:\Users\jeNguyen\Desktop\reconciled_vendors_emailing list.csv"

$sig = "Jennifer Nguyen`n`nService Accounting Representative`nProfessional Coffee Machines – The Americas`n`nSEB Professional North America | 15501 Red Hill Ave. STE 200 | Tustin, CA 92780 | USA`n`nJeNGUYEN@seb-professional.com | www.wmfnorthamerica.com | www.schaererusa.com"

$data = Import-Csv $path 

Start-Process Outlook
$o = New-Object -com Outlook.Application

for($i = 0; $i -lt $data.Length; $i++) {
    
    $Name = $data.Vendor[$i].Trim();
    $Email = $data.Email[$i];
    $Date = $data.Date[$i];
    $mail = $o.CreateItem(0)
    $mail.subject = "Confirmation of 2019 reconciliation"
    $mail.body = "Hello "+ $Name+ "`, `n`nThis is to confirm that we have now completed our 2019 reconciliation between Schaerer USA (SEB PRO) and "+$Name+ " as of "+$Date+ "`.`nIf there are any invoices dated 2019 or prior to 2019 remaining on your statement, please review and remove them from your statement since our records indicate that we have now paid all 2019 and prior past due balances.  Please also reply to this email with your latest statement to confirm receipt of this email.`n`nIf you have any questions or concerns, please respond to this email and we be in contact with you.`n`nThank you for your efforts.`n"
    $mail.body +=$sig

    $mail.To = $Email
    $mail.Send()

}