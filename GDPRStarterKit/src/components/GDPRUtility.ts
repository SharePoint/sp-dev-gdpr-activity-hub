export class GDPRUtility {

  public static getInitials(fullname: string): string {

    if (!fullname) {
        return (null);
    }
    
    var parts = fullname.split(' ');
    
    var initials = "";
    parts.forEach(p => {
        if (p.length > 0)
        {
            initials = initials.concat(p.substring(0, 1).toUpperCase());
        }
    });

    return (initials);
  }

  public static getPersonaImage(siteUrl: string, accountName: string): string {
    let tenantBaseUrl: string = siteUrl.substring(0, siteUrl.indexOf("sharepoint.com") + 14);
    let imageBaseUrl =  tenantBaseUrl + "/_layouts/15/userphoto.aspx?size=S&accountname=";

    if (accountName.indexOf("|membership|") > 0) {
      accountName = accountName.substring(accountName.indexOf("|membership|") + 12);
    }

    return (imageBaseUrl + accountName);
  }
}