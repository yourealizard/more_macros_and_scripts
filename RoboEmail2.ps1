$path = "C:\Users\jeNguyen\Desktop\Second Round, no response.csv"
$file = "C:\Users\jeNguyen\Desktop\VENDORRECONCILIATIONLETTER.docx"

$sig = "Jennifer Nguyen`n`nService Accounting Representative`nProfessional Coffee Machines – The Americas`n`nSEB Professional North America | 15501 Red Hill Ave. STE 200 | Tustin, CA 92780 | USA`n`nJeNGUYEN@seb-professional.com | www.wmfnorthamerica.com | www.schaererusa.com"

$data = Import-Csv $path 

Start-Process Outlook
$o = New-Object -com Outlook.Application

for($i = 0; $i -lt $data.Length; $i++) {
    
    $Name = $data.Vendor[$i].Trim();
    $Email = $data.Email[$i];
    $Date = $data.Date[$i];
    $mail = $o.CreateItem(0)
    $mail.subject = "2019 Outstanding Invoices"
    $mail.body = "Dear Vendor`:`n`n`nSEB Professional is conducting an audit to ensure that we have completed payment to our vendors for all outstanding invoices for the year of 2019`.  We respectfully request that you send an updated statement of your account or send an email indicating that you have no outstanding 2019 invoices to billing-us@seb-professional.com by August 1`, 2020`.  If we have not received communication from you by August 1`, 2020 we will consider that there is not an outstanding balance owed to you by SEB Professional`.  Invoices brought to our attention after that day may not be considered for payment`.  If you have any questions`, please feel free to contact us at the email listed above`.`n`nBest Regards`,`n`nSEB Professional"
    $mail.Attachments.Add($file)
    $mail.To = $Email
    $mail.Send()

}