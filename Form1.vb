Imports System.IO
Imports System.Net.Http
Imports HtmlAgilityPack
Imports Newtonsoft.Json
Imports System.Data.SQLite
Imports System.Security.Policy
Imports System.Transactions

Public Class Form1
    Private connStr As String = "Data Source=E:\sqlite\FirstTable.db;Version=3;"
    Private conn As SQLiteConnection
    Private transaction As SQLiteTransaction

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GetCompanyDetails()


        '1. Input  - Letter A,B,C,D,E,F and so on
        '2. Open coresponding file (A.txt, B.txt, C.txt - any one file as per input)
        '3. Read each line from the file and get the href (pagenum, linenum, href)
        '4. Get company id from href
        '5. Replace this company id to discovery url
        '6. Send http request to this new url & get returned json file
        '7. Parse the json file. to get each field.
        '8. Store each field to SQL
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        '1. Input - Letter A, B, C, D, E, F, etc.

        ' Replace with your input logic
    End Sub
    Private Function ConvertStringToDecimal(value As String) As Decimal?
        If String.IsNullOrWhiteSpace(value) Then
            Return Nothing
        End If

        Dim cleanedValue As String = value.Trim()

        ' Check if the string ends with "K" and remove it
        If cleanedValue.EndsWith("K", StringComparison.OrdinalIgnoreCase) Then
            cleanedValue = cleanedValue.Substring(0, cleanedValue.Length - 1)

            ' Attempt to parse the cleaned value as a decimal, and multiply by 1000
            Dim result As Decimal
            If Decimal.TryParse(cleanedValue, result) Then
                Return result * 1000
            Else
                ' Parsing failed, return null or handle the error as needed
                Return Nothing
            End If
        End If

        ' If the string doesn't end with "K", try to parse it directly as a decimal
        Dim directResult As Decimal
        If Decimal.TryParse(cleanedValue, directResult) Then
            Return directResult
        Else
            ' Parsing failed, return null or handle the error as needed
            Return Nothing
        End If
    End Function

    ' ...





    ' ...

    Private Sub GetCompanyDetails()

        Dim inputLetter As String = TextBox1.Text

        '2. Open corresponding file (A.txt, B.txt, C.txt - any one file as per input)
        Dim filePath As String = Path.Combine(Application.StartupPath, $"{inputLetter}.txt")

        If File.Exists(filePath) Then
            Try
                '3. Read each line from the file and get the href (pagenum, linenum, href)
                Dim lines As String() = File.ReadAllLines(filePath)

                Dim RCount As Double
                For Each line As String In lines
                    Dim parts As String() = line.Split(",")
                    Dim pageNumber As Integer = Integer.Parse(parts(0))
                    Dim lineNum As Integer = Integer.Parse(parts(1))
                    Dim href As String = parts(2).Trim()
                    RCount += 1
                    Label1.Text = RCount
                    '4. Get company id from href

                    Dim companyId As String = href.Substring(href.LastIndexOf("/") + 1)
                    Dim sAPIURL As String = "https://discovery-api.apollo.io/api/v1/discovery/modality/organizations/entity/{companyId}"
                    Dim apiUrl As String = sAPIURL & companyId


                    '5. Replace this company id in the discovery URL
                    Dim discoverUrl As String = sAPIURL.Replace("{companyId}", companyId)

                    '6. Send HTTP request to this new URL & get the returned JSON file
                    Dim jsonData As String = FetchDataFromApi(discoverUrl)
                    Debug.Print("")
                    Debug.Print("")
                    Debug.Print("")
                    Debug.Print("---------------------------")
                    'jsonData = jsonData.Replace("[]", "null")
                    Debug.Print(jsonData)

                    '7. Parse the JSON file to get each field
                    Dim companyDetails As CompanyDetails = JsonConvert.DeserializeObject(Of CompanyDetails)(jsonData)

                    '8. Store each field in SQL
                    InsertDataIntoSQLite(companyDetails)
                    ' Print a message to indicate progress
                    Console.WriteLine($"Processed companyId: {companyId}")

                    ' Add a delay (optional) to avoid overloading the server with requests
                    Threading.Thread.Sleep(1000) ' Sleep for 1 second

                Next

            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
            End Try
        Else
            MessageBox.Show($"The '{inputLetter}.txt' file does not exist in the application's directory.")
        End If

        MessageBox.Show("Done")
    End Sub
    Private Class CompanyDetails
        Public Property Id As String
        Public Property Name As String
        Public Property LogoUrl As String
        Public Property WebsiteUrl As String
        Public Property Location As LocationDetails
        Public Property SocialLinks As SocialLinksDetails
        Public Property PhoneNumber As String
        Public Property EmployeeCount As Integer
        Public Property JobPostingUrls As List(Of String)
        Public Property Technologies As List(Of Technology)
        Public Property RetailLocationCount As Integer
        Public Property Revenue As String
        Public Property Industries As List(Of String)
        Public Property Keywords As List(Of String)
        Public Property Description As String
        Public Property LastUpdated As DateTime
        Public Property IsPublic As Boolean
        Public Property PubliclyTradedSymbol As String
        Public Property PubliclyTradedExchange As String
        Public Property MarketCap As Decimal?
        Public Property TotalFundingAmount As Decimal?
        Public Property LatestFundingAmount As Decimal?
        Public Property FoundedAt As Date
        Public Property SimilarCompanies As List(Of String)
    End Class

    Private Class LocationDetails
        Public Property StreetAddress As String
        Public Property City As String
        Public Property State As String
        Public Property PostalCode As String
        Public Property Country As String
    End Class

    Private Class SocialLinksDetails
        Public Property TwitterUrl As String
        Public Property FacebookUrl As String
        Public Property InstagramUrl As String
        Public Property LinkedinUrl As String
        Public Property YoutubeUrl As String
        Public Property AngellistUrl As String
        Public Property CrunchbaseUrl As String
        Public Property BlogUrl As String
    End Class

    Private Class Technology
        Public Property Name As String
        Public Property Category As String
    End Class

    Private Function FetchDataFromApi(apiUrl As String) As String
        Using client As New HttpClient()
            Dim response As HttpResponseMessage = client.GetAsync(apiUrl).Result


            If response.IsSuccessStatusCode Then
                Return response.Content.ReadAsStringAsync().Result
            Else
                Throw New Exception("Failed to fetch data from the API. Status Code: " & response.StatusCode)
            End If
        End Using
    End Function
    Private Sub InsertDataIntoSQLite(companyDetails As CompanyDetails)
        Using conn As New SQLiteConnection(connStr)
            conn.Open()
            transaction = conn.BeginTransaction()
            ' Check if a record with the same id already exists
            Dim idExists As Boolean
            Using checkCmd As New SQLiteCommand(conn)
                checkCmd.CommandText = "SELECT COUNT(*) FROM ZCompanyDetails WHERE id = @id"
                checkCmd.Parameters.AddWithValue("@id", companyDetails.Id)
                idExists = CInt(checkCmd.ExecuteScalar()) > 0
            End Using
            If Not idExists Then

                Using cmd As New SQLiteCommand(conn)
                    cmd.CommandText = "INSERT INTO ZCompanyDetails (id, name, logo_url, website_url, street_address, city, state, postal_code, country, twitter_url, facebook_url, instagram_url, linkedin_url, youtube_url, angellist_url, crunchbase_url, blog_url, phone_number, employee_count, retail_location_count, revenue, description, last_updated, is_public, publicly_traded_symbol, publicly_traded_exchange, market_cap, total_funding_amount, latest_funding_amount, founded_at) " &
                                  "VALUES (@id, @name, @logo_url, @website_url, @street_address, @city, @state, @postal_code, @country, @twitter_url, @facebook_url, @instagram_url, @linkedin_url, @youtube_url, @angellist_url, @crunchbase_url, @blog_url, @phone_number, @employee_count, @retail_location_count, @revenue, @description, @last_updated, @is_public, @publicly_traded_symbol, @publicly_traded_exchange, @market_cap, @total_funding_amount, @latest_funding_amount, @founded_at)"

                    cmd.Parameters.AddWithValue("@id", companyDetails.Id)
                    cmd.Parameters.AddWithValue("@name", companyDetails.Name)
                    cmd.Parameters.AddWithValue("@logo_url", If(companyDetails.LogoUrl, DBNull.Value))
                    cmd.Parameters.AddWithValue("@website_url", If(companyDetails.WebsiteUrl, DBNull.Value))
                    cmd.Parameters.AddWithValue("@street_address", If(companyDetails.Location?.StreetAddress, DBNull.Value))
                    cmd.Parameters.AddWithValue("@city", If(companyDetails.Location?.City, DBNull.Value))
                    cmd.Parameters.AddWithValue("@state", If(companyDetails.Location?.State, DBNull.Value))
                    cmd.Parameters.AddWithValue("@postal_code", If(companyDetails.Location?.PostalCode, DBNull.Value))
                    cmd.Parameters.AddWithValue("@country", If(companyDetails.Location?.Country, DBNull.Value))

                    ' Check if SocialLinks is not null before accessing its properties
                    If companyDetails.SocialLinks IsNot Nothing Then
                        cmd.Parameters.AddWithValue("@twitter_url", companyDetails.SocialLinks.TwitterUrl)
                        cmd.Parameters.AddWithValue("@facebook_url", companyDetails.SocialLinks.FacebookUrl)
                        cmd.Parameters.AddWithValue("@instagram_url", companyDetails.SocialLinks.InstagramUrl)
                        cmd.Parameters.AddWithValue("@linkedin_url", companyDetails.SocialLinks.LinkedinUrl)
                        cmd.Parameters.AddWithValue("@youtube_url", companyDetails.SocialLinks.YoutubeUrl)
                        cmd.Parameters.AddWithValue("@angellist_url", companyDetails.SocialLinks.AngellistUrl)
                        cmd.Parameters.AddWithValue("@crunchbase_url", companyDetails.SocialLinks.CrunchbaseUrl)
                        cmd.Parameters.AddWithValue("@blog_url", companyDetails.SocialLinks.BlogUrl)
                    Else
                        ' If SocialLinks is null, set these parameters to NULL or empty values as needed
                        cmd.Parameters.AddWithValue("@twitter_url", DBNull.Value)
                        cmd.Parameters.AddWithValue("@facebook_url", DBNull.Value)
                        cmd.Parameters.AddWithValue("@instagram_url", DBNull.Value)
                        cmd.Parameters.AddWithValue("@linkedin_url", DBNull.Value)
                        cmd.Parameters.AddWithValue("@youtube_url", DBNull.Value)
                        cmd.Parameters.AddWithValue("@angellist_url", DBNull.Value)
                        cmd.Parameters.AddWithValue("@crunchbase_url", DBNull.Value)
                        cmd.Parameters.AddWithValue("@blog_url", DBNull.Value)
                    End If

                    cmd.Parameters.AddWithValue("@phone_number", companyDetails.PhoneNumber)
                    cmd.Parameters.AddWithValue("@employee_count", companyDetails.EmployeeCount)
                    cmd.Parameters.AddWithValue("@retail_location_count", companyDetails.RetailLocationCount)
                    cmd.Parameters.AddWithValue("@revenue", companyDetails.Revenue)
                    cmd.Parameters.AddWithValue("@description", companyDetails.Description)
                    cmd.Parameters.AddWithValue("@last_updated", companyDetails.LastUpdated)
                    cmd.Parameters.AddWithValue("@is_public", companyDetails.IsPublic)
                    cmd.Parameters.AddWithValue("@publicly_traded_symbol", companyDetails.PubliclyTradedSymbol)
                    cmd.Parameters.AddWithValue("@publicly_traded_exchange", companyDetails.PubliclyTradedExchange)
                    cmd.Parameters.AddWithValue("@market_cap", companyDetails.MarketCap)
                    cmd.Parameters.AddWithValue("@total_funding_amount", companyDetails.TotalFundingAmount)
                    cmd.Parameters.AddWithValue("@latest_funding_amount", companyDetails.LatestFundingAmount)
                    cmd.Parameters.AddWithValue("@founded_at", companyDetails.FoundedAt)

                    cmd.ExecuteNonQuery()
                End Using
            Else
                Console.WriteLine($"Record with id '{companyDetails.Id}' already exists. Skipping insertion.")
            End If
            transaction.Commit()

            conn.Close()


        End Using
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
    ' Define CompanyDetails class to match the structure of the JSON response




    ' Implement the FetchDataFromApi and InsertDataIntoSQLite functions as in your previous code
    ' ...

    ' Other methods and event handlers...
End Class


