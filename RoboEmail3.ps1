$path = "C:\Users\jeNguyen\Desktop\contact6.11.21.csv"

$sig = "Jennifer Nguyen`n`nService Accounting Representative`nProfessional Coffee Machines – The Americas`n`nSEB Professional North America | 15501 Red Hill Ave. STE 200 | Tustin, CA 92780 | USA`n`nJeNGUYEN@seb-professional.com | www.wmfnorthamerica.com | www.schaererusa.com"

$data = Import-Csv $path 

Start-Process Outlook
$o = New-Object -com Outlook.Application

for($i = 0; $i -lt $data.Length; $i++) {
    
    $Name = $data.Vendor[$i].Trim();
    $Email = $data.Email[$i];
    $Date = $data.Date[$i];
    $mail = $o.CreateItem(0)
    $mail.subject = "Statement Request"
    $mail.body = "Hello "+ $Name+ "`, `n`nI am currently doing account reconciliations. We want to have our accounts up to date and we need your help to have accurate information.  Can we please get a copy of the most recent statement?  The last statement I have is from $Date.  Thank you for your effort!`n"
    $mail.body +=$sig
    $mail.To = $Email
    #$mail.cc = "ENAVARRO@seb-professional.com"
    $mail.Send()

}