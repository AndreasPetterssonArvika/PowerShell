function Confirm-ADUser
{
    param(
        $Email
    )

    # Confirm that email-adress exists in Active Directory
    Get-ADUser

}