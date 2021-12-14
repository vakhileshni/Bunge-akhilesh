$svo=sfdx force:data:soql:query -q "Select COUNTRY__C,EMAIL,ID,ISACTIVE,MANAGERID,NAME,PROFILEID,USERROLEID,FEDERATIONIDENTIFIER FROM User" -u -Pro -r csv
$svo|out-file atten.csv