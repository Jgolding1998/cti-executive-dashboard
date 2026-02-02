// SyteLine IDO API Connection for Power BI
// This query connects directly to SyteLine REST API

let
    // Configuration
    BaseUrl = "https://csi10g.erpsl.inforcloudsuite.com/IDORequestService/ido",
    Tenant = "GVNDYXUFKHB5VMB6_PRD_CTI",
    Username = "gary.phillips@godlan.com",
    Password = "Crwthtithing2$",
    
    // Get Token
    TokenUrl = BaseUrl & "/token/" & Tenant & "/" & Username & "/" & Uri.EscapeDataString(Password),
    TokenResponse = Json.Document(Web.Contents(TokenUrl)),
    Token = TokenResponse[Token],
    
    // Function to call SyteLine API
    GetIDOData = (ido as text, properties as text, optional filter as text, optional recordCap as number) as table =>
        let
            cap = if recordCap = null then 5000 else recordCap,
            filterParam = if filter = null then "" else "&filter=" & Uri.EscapeDataString(filter),
            url = BaseUrl & "/load/" & ido & "?properties=" & properties & "&recordcap=" & Text.From(cap) & filterParam,
            response = Json.Document(Web.Contents(url, [Headers=[Authorization=Token]])),
            items = response[Items],
            result = if List.Count(items) > 0 then Table.FromRecords(items) else #table({}, {})
        in
            result
in
    GetIDOData
