<%
Option Explicit

' --- Bond Data ---
Dim bonds(5, 3)
bonds(0, 0) = "Bond 1YR - ₦1,000 - 8.5% - 1 Year": bonds(0, 1) = 1000: bonds(0, 2) = 8.5: bonds(0, 3) = 1
bonds(1, 0) = "Bond 3YR - ₦2,500 - 10.2% - 3 Years": bonds(1, 1) = 2500: bonds(1, 2) = 10.2: bonds(1, 3) = 3
bonds(2, 0) = "Bond 5YR - ₦5,000 - 12.0% - 5 Years": bonds(2, 1) = 5000: bonds(2, 2) = 12.0: bonds(2, 3) = 5
bonds(3, 0) = "Bond 10YR - ₦10,000 - 13.5% - 10 Years": bonds(3, 1) = 10000: bonds(3, 2) = 13.5: bonds(3, 3) = 10
bonds(4, 0) = "Bond 15YR - ₦15,000 - 14.8% - 15 Years": bonds(4, 1) = 15000: bonds(4, 2) = 14.8: bonds(4, 3) = 15
bonds(5, 0) = "Bond 20YR - ₦25,000 - 15.5% - 20 Years": bonds(5, 1) = 25000: bonds(5, 2) = 15.5: bonds(5, 3) = 20

' --- Defaults ---
Dim selectedBond, quantity, monthlyContribution, reinvestInterest, compoundingFreq
selectedBond = 0
quantity = 1
monthlyContribution = 0
reinvestInterest = False
compoundingFreq = 4

' --- Results ---
Dim bondName, bondPrice, bondYield, yieldRate, years, consideration, commission, fees, finalAmount, totalReturn
Dim projections, year, yearlyAmount, monthlyContribTotal, totalInterest

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Request.Form("bondSelect") <> "" Then selectedBond = CInt(Request.Form("bondSelect"))
    If Request.Form("quantity") <> "" Then quantity = CInt(Request.Form("quantity"))
    If Request.Form("monthlyContribution") <> "" Then monthlyContribution = CDbl(Request.Form("monthlyContribution"))
    If Request.Form("reinvestInterest") = "on" Then reinvestInterest = True Else reinvestInterest = False
    If Request.Form("compoundingFreq") <> "" Then compoundingFreq = CInt(Request.Form("compoundingFreq"))
    
    bondName = bonds(selectedBond, 0)
    bondPrice = bonds(selectedBond, 1)
    bondYield = bonds(selectedBond, 2)
    yieldRate = bondYield / 100
    years = bonds(selectedBond, 3)
    consideration = bondPrice * quantity
    commission = consideration * 0.0075
    fees = consideration * 0.002

    Dim principal, rate, n, t
    principal = consideration
    rate = yieldRate
    n = compoundingFreq
    t = years

    Set projections = Server.CreateObject("Scripting.Dictionary")
    For year = 0 To years
        If year = 0 Then
            yearlyAmount = principal
            monthlyContribTotal = 0
            totalInterest = 0
        Else
            yearlyAmount = principal * ((1 + rate / n) ^ (n * year))
            If monthlyContribution > 0 Then
                Dim monthlyRate, months, futureValueAnnuity
                monthlyRate = rate / 12
                months = year * 12
                If monthlyRate > 0 Then
                    futureValueAnnuity = monthlyContribution * ((1 + monthlyRate) ^ months - 1) / monthlyRate
                Else
                    futureValueAnnuity = monthlyContribution * months
                End If
                yearlyAmount = yearlyAmount + futureValueAnnuity
                monthlyContribTotal = monthlyContribution * 12 * year
            End If
            totalInterest = yearlyAmount - principal - monthlyContribTotal
        End If
        projections.Add year, Array(yearlyAmount, totalInterest)
    Next
    finalAmount = yearlyAmount
    totalReturn = finalAmount - consideration
End If

Function FormatCurrency(amount)
    FormatCurrency = "₦" & FormatNumber(amount, 2)
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8" />
    <title>Bond Investment Calculator</title>
    <style>
        body {
            font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
            background-color: #1a1a1a;
            color: #ffffff;
            min-height: 100vh;
            margin: 0;
        }
        .header {
            background-color: #2d2d2d;
            padding: 15px 20px;
            border-bottom: 1px solid #444;
            display: flex;
            justify-content: flex-start;
            align-items: center;
            gap: 12px;
        }
        .logo-container {
          display:flex;
          align-items:center;
          gap: 8px;
        }

        .logo-img{
          height: 32px;
          width: 32px;
          object-fit:contain;
          border-radius: 6px;
          background: #fff;
        }


        .header h1 {
            font-size: 18px;
            font-weight: 600;
            margin: 0;
        }
        .main-container {
            display: grid;
            grid-template-columns: 1fr 2fr;
            gap: 10px;
            padding: 30px 10px 10px 10px;
        }
        .left-panel {
            background-color: #252525;
            border-radius: 8px;
            padding: 20px;
            border: 1px solid #333;
            min-width: 320px;
            max-width: 400px;
        }
        .left-panel label {
            display: block;
            margin-bottom: 6px;
            font-size: 14px;
            font-weight: 500;
        }
        .left-panel select,
        .left-panel input[type="number"],
        .left-panel input[type="text"] {
            width: 100%;
            padding: 8px;
            margin-bottom: 14px;
            border-radius: 4px;
            border: 1px solid #444;
            background: #181818;
            color: #fff;
            font-size: 15px;
        }
        .left-panel input[type="checkbox"] {
            margin-right: 6px;
        }
        .btn-submit {
            width: 100%;
            background: #0078ff;
            color: #fff;
            border: none;
            padding: 12px;
            border-radius: 4px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            margin-top: 10px;
        }
        .btn-submit:hover {
            background: #005fcc;
        }
        .calculation-results {
            margin-top: 20px;
        }
        .result-row {
            display: flex;
            justify-content: space-between;
            margin-bottom: 8px;
            font-size: 15px;
        }
        .result-row .naira {
            color: #3ecf8e;
            font-weight: 600;
        }
        .result-row:last-child {
            font-size: 16px;
            font-weight: 600;
        }
        .right-panel {
            background-color: #252525;
            border-radius: 8px;
            padding: 20px;
            border: 1px solid #333;
            min-width: 400px;
        }
        .panel-section {
            margin-bottom: 24px;
        }
        .section-title {
            font-size: 17px;
            font-weight: 600;
            margin-bottom: 12px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            background: #181818;
            margin-bottom: 18px;
        }
        th, td {
            padding: 8px 10px;
            border-bottom: 1px solid #333;
            text-align: right;
        }
        th {
            background: #232323;
            color: #3ecf8e;
            font-weight: 600;
            text-align: right;
        }
        td:first-child, th:first-child {
            text-align: left;
        }
        .market-info-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 10px;
        }
        .market-metric {
            background: #181818;
            border-radius: 4px;
            padding: 10px 8px;
            text-align: center;
        }
        .metric-label {
            font-size: 12px;
            color: #aaa;
        }
        .metric-value {
            font-size: 16px;
            font-weight: 600;
            color: #3ecf8e;
        }
        @media (max-width: 900px) {
            .main-container {
                grid-template-columns: 1fr;
            }
            .right-panel {
                min-width: unset;
            }
        }
    </style>
</head>
<body>
    <div class="header">
      <div class="logo-container">
        <img src="./logo.png" alt="Compoundly Logo" class="logo-img">
      </div>
        <h1>Compoundly</h1>
    </div>
    <form method="post" action="index.asp">
        <div class="main-container">
            <div class="left-panel">
                <label for="bondSelect">Select Bond</label>
                <select name="bondSelect" id="bondSelect">
                    <% Dim i
                    For i = 0 To UBound(bonds, 1)
                        Dim selected
                        If i = selectedBond Then selected = "selected" Else selected = ""
                    %>
                        <option value="<%=i%>" <%=selected%>><%=bonds(i,0)%></option>
                    <% Next %>
                </select>
                <label for="quantity">Quantity</label>
                <input type="number" name="quantity" id="quantity" min="1" value="<%=quantity%>" required />

                <label for="monthlyContribution">₦ Monthly Additional Investment (Optional)</label>
                <input type="number" name="monthlyContribution" id="monthlyContribution" min="0" value="<%=monthlyContribution%>" />

                <label>
                    <input type="checkbox" name="reinvestInterest" <% If reinvestInterest Then Response.Write("checked") %> />
                    Reinvest Interest
                </label>

                <label for="compoundingFreq">Compounding Frequency</label>
                <select name="compoundingFreq" id="compoundingFreq">
                    <option value="1" <% If compoundingFreq=1 Then Response.Write("selected") %>>Annually (1 time/year)</option>
                    <option value="2" <% If compoundingFreq=2 Then Response.Write("selected") %>>Semi-Annually (2 times/year)</option>
                    <option value="4" <% If compoundingFreq=4 Then Response.Write("selected") %>>Quarterly (4 times/year)</option>
                    <option value="12" <% If compoundingFreq=12 Then Response.Write("selected") %>>Monthly (12 times/year)</option>
                </select>
                <button class="btn-submit" type="submit">Calculate Investment</button>
                <% If Request.ServerVariables("REQUEST_METHOD") = "POST" Then %>
                <div class="calculation-results">
                    <div class="result-row">
                        <span>Consideration</span>
                        <span class="naira"><%=FormatCurrency(consideration)%></span>
                    </div>
                    <div class="result-row">
                        <span>Commission</span>
                        <span class="naira"><%=FormatCurrency(commission)%></span>
                    </div>
                    <div class="result-row">
                        <span>Fees</span>
                        <span class="naira"><%=FormatCurrency(fees)%></span>
                    </div>
                    <div class="result-row">
                        <span>Estimated Total Return</span>
                        <span class="naira"><%=FormatCurrency(finalAmount)%></span>
                    </div>
                </div>
                <% End If %>
            </div>
            <div class="right-panel">
                <div class="panel-section">
                    <div class="section-title">Bond Investment Growth Over Time</div>
                    <% If Request.ServerVariables("REQUEST_METHOD") = "POST" Then %>
                        <table>
                            <tr>
                                <th>Year</th>
                                <th>Total Investment Value</th>
                                <th>Interest Earned</th>
                            </tr>
                            <% For year = 0 To years
                                Dim proj
                                proj = projections(year)
                            %>
                            <tr>
                                <td><%=year%></td>
                                <td><%=FormatCurrency(proj(0))%></td>
                                <td><%=FormatCurrency(proj(1))%></td>
                            </tr>
                            <% Next %>
                        </table>
                    <% Else %>
                        <div style="color:#aaa;">Fill the form and click Calculate Investment to see projections.</div>
                    <% End If %>
                </div>
                <div class="panel-section">
                    <div class="section-title">Market Information</div>
                    <div class="market-info-grid">
                        <div class="market-metric">
                            <div class="metric-label">LAST PRICE</div>
                            <div class="metric-value"><% If Request.ServerVariables("REQUEST_METHOD") = "POST" Then Response.Write(bondPrice) Else Response.Write("-") %></div>
                        </div>
                        <div class="market-metric">
                            <div class="metric-label">CHANGE</div>
                            <div class="metric-value">+0.00</div>
                        </div>
                        <div class="market-metric">
                            <div class="metric-label">% CHANGE</div>
                            <div class="metric-value">0.00%</div>
                        </div>
                        <div class="market-metric">
                            <div class="metric-label">YIELD</div>
                            <div class="metric-value"><% If Request.ServerVariables("REQUEST_METHOD") = "POST" Then Response.Write(bondYield & "%") Else Response.Write("-") %></div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </form>
</body>
</html>
