$cn = [ADSI]"WinNT://localhost"
$user = $cn.Create("User","DesireLocalAdminAccountName")
$user.SetPassword("DesiredPassword")
$user.SetInfo()
$user.Description = "Local Admin Account"
$user.SetInfo()