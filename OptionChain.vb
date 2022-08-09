Public Class OptionChain
    Public Function ImpliedCallvolatility(ByVal underlyingPrice As Double, ByVal exercisePrice As Double, ByVal daysToExpiry As Double, ByVal interest As Double, ByVal premium As Double, ByVal dividendYield As Double) As Double
        If Not underlyingPrice > 0 Then Return 0
        If Not exercisePrice > 0 Then Return 0
        If Not daysToExpiry > 0 Then Return 0
        If Not interest > -1 Then Return 0
        If Not premium > 0 Then Return 0

        Dim high As Double = 5
        Dim low As Double = 0
        Do While (high - low) > 0.0001
            If CallOption(underlyingPrice, exercisePrice, daysToExpiry, interest, ((high + low) / 2) * 100, dividendYield) > premium Then
                high = (high + low) / 2
            Else
                low = (high + low) / 2
            End If
        Loop
        Return ((high + low) / 2) * 100
    End Function

    Public Function ImpliedPutvolatility(ByVal underlyingPrice As Double, ByVal exercisePrice As Double, ByVal daysToExpiry As Double, ByVal interest As Double, ByVal premium As Double, ByVal dividendYield As Double) As Double
        If Not underlyingPrice > 0 Then Return 0
        If Not exercisePrice > 0 Then Return 0
        If Not daysToExpiry > 0 Then Return 0
        If Not interest > -1 Then Return 0
        If Not premium > 0 Then Return 0

        Dim high As Double = 5
        Dim low As Double = 0
        Do While (high - low) > 0.0001
            If PutOption(underlyingPrice, exercisePrice, daysToExpiry, interest, ((high + low) / 2) * 100, dividendYield) > premium Then
                high = (high + low) / 2
            Else
                low = (high + low) / 2
            End If
        Loop
        Return ((high + low) / 2) * 100
    End Function

    Public Function CallOption(ByVal underlyingPrice As Double, ByVal exercisePrice As Double, ByVal daysToExpiry As Double, ByVal interest As Double, ByVal volatility As Double, ByVal dividendYield As Double) As Double
        If Not underlyingPrice > 0 Then Return 0
        If Not exercisePrice > 0 Then Return 0
        If Not daysToExpiry > 0 Then Return 0
        If Not interest > -1 Then Return 0
        If Not volatility > 0 Then Return 0

        'Convert Days to Years
        Dim yearsToExpiry As Double = daysToExpiry / 365

        'Convert Numbers to %
        interest = interest / 100
        volatility = volatility / 100
        dividendYield = dividendYield / 100

        Return Math.Exp(-dividendYield * yearsToExpiry) * underlyingPrice * NormSDist(dOne(underlyingPrice, exercisePrice, yearsToExpiry, interest, volatility, dividendYield)) - exercisePrice * Math.Exp(-interest * yearsToExpiry) * NormSDist(dOne(underlyingPrice, exercisePrice, yearsToExpiry, interest, volatility, dividendYield) - volatility * Math.Sqrt(yearsToExpiry))
    End Function

    Public Function PutOption(ByVal underlyingPrice As Double, ByVal exercisePrice As Double, ByVal daysToExpiry As Double, ByVal interest As Double, ByVal volatility As Double, ByVal dividendYield As Double) As Double
        If Not underlyingPrice > 0 Then Return 0
        If Not exercisePrice > 0 Then Return 0
        If Not daysToExpiry > 0 Then Return 0
        If Not interest > -1 Then Return 0
        If Not volatility > 0 Then Return 0

        'Convert Days to Years
        Dim yearsToExpiry As Double = daysToExpiry / 365

        'Convert Numbers to %
        interest = interest / 100
        volatility = volatility / 100
        dividendYield = dividendYield / 100

        Return exercisePrice * Math.Exp(-interest * yearsToExpiry) * NormSDist(-dTwo(underlyingPrice, exercisePrice, yearsToExpiry, interest, volatility, dividendYield)) - Math.Exp(-dividendYield * yearsToExpiry) * underlyingPrice * NormSDist(-dOne(underlyingPrice, exercisePrice, yearsToExpiry, interest, volatility, dividendYield))
    End Function

    Public Function CallDelta(ByVal underlyingPrice As Double, ByVal exercisePrice As Double, ByVal daysToExpiry As Double, ByVal interest As Double, ByVal volatility As Double, ByVal dividendYield As Double) As Double
        If Not underlyingPrice > 0 Then Return 0
        If Not exercisePrice > 0 Then Return 0
        If Not daysToExpiry > 0 Then Return 0
        If Not interest > -1 Then Return 0
        If Not volatility > 0 Then Return 0

        'Convert Days to Years
        Dim yearsToExpiry As Double = daysToExpiry / 365

        'Convert Numbers to %
        interest = interest / 100
        volatility = volatility / 100
        dividendYield = dividendYield / 100

        Return NormSDist(dOne(underlyingPrice, exercisePrice, yearsToExpiry, interest, volatility, dividendYield))
    End Function

    Public Function PutDelta(ByVal underlyingPrice As Double, ByVal exercisePrice As Double, ByVal daysToExpiry As Double, ByVal interest As Double, ByVal volatility As Double, ByVal dividendYield As Double) As Double
        If Not underlyingPrice > 0 Then Return 0
        If Not exercisePrice > 0 Then Return 0
        If Not daysToExpiry > 0 Then Return 0
        If Not interest > -1 Then Return 0
        If Not volatility > 0 Then Return 0

        'Convert Days to Years
        Dim yearsToExpiry As Double = daysToExpiry / 365

        'Convert Numbers to %
        interest = interest / 100
        volatility = volatility / 100
        dividendYield = dividendYield / 100

        Return NormSDist(dOne(underlyingPrice, exercisePrice, yearsToExpiry, interest, volatility, dividendYield)) - 1
    End Function

    Public Function CallTheta(ByVal underlyingPrice As Double, ByVal exercisePrice As Double, ByVal daysToExpiry As Double, ByVal interest As Double, ByVal volatility As Double, ByVal dividendYield As Double) As Double
        If Not underlyingPrice > 0 Then Return 0
        If Not exercisePrice > 0 Then Return 0
        If Not daysToExpiry > 0 Then Return 0
        If Not interest > -1 Then Return 0
        If Not volatility > 0 Then Return 0

        'Convert Days to Years
        Dim yearsToExpiry As Double = daysToExpiry / 365

        'Convert Numbers to %
        interest = interest / 100
        volatility = volatility / 100
        dividendYield = dividendYield / 100

        Dim ct As Double = -(underlyingPrice * volatility * NdOne(underlyingPrice, exercisePrice, yearsToExpiry, interest, volatility, dividendYield)) / (2 * Math.Sqrt(yearsToExpiry)) - interest * exercisePrice * Math.Exp(-interest * (yearsToExpiry)) * NdTwo(underlyingPrice, exercisePrice, yearsToExpiry, interest, volatility, dividendYield)
        Return ct / 365
    End Function

    Public Function PutTheta(ByVal underlyingPrice As Double, ByVal exercisePrice As Double, ByVal daysToExpiry As Double, ByVal interest As Double, ByVal volatility As Double, ByVal dividendYield As Double) As Double
        If Not underlyingPrice > 0 Then Return 0
        If Not exercisePrice > 0 Then Return 0
        If Not daysToExpiry > 0 Then Return 0
        If Not interest > -1 Then Return 0
        If Not volatility > 0 Then Return 0

        'Convert Days to Years
        Dim yearsToExpiry As Double = daysToExpiry / 365

        'Convert Numbers to %
        interest = interest / 100
        volatility = volatility / 100
        dividendYield = dividendYield / 100

        Dim pt As Double = -(underlyingPrice * volatility * NdOne(underlyingPrice, exercisePrice, yearsToExpiry, interest, volatility, dividendYield)) / (2 * Math.Sqrt(yearsToExpiry)) + interest * exercisePrice * Math.Exp(-interest * (yearsToExpiry)) * (1 - NdTwo(underlyingPrice, exercisePrice, yearsToExpiry, interest, volatility, dividendYield))
        Return pt / 365
    End Function

    Public Function CallRho(ByVal underlyingPrice As Double, ByVal exercisePrice As Double, ByVal daysToExpiry As Double, ByVal interest As Double, ByVal volatility As Double, ByVal dividendYield As Double) As Double
        If Not underlyingPrice > 0 Then Return 0
        If Not exercisePrice > 0 Then Return 0
        If Not daysToExpiry > 0 Then Return 0
        If Not interest > -1 Then Return 0
        If Not volatility > 0 Then Return 0

        'Convert Days to Years
        Dim yearsToExpiry As Double = daysToExpiry / 365

        'Convert Numbers to %
        interest = interest / 100
        volatility = volatility / 100
        dividendYield = dividendYield / 100

        Return 0.01 * exercisePrice * yearsToExpiry * Math.Exp(-interest * yearsToExpiry) * NormSDist(dTwo(underlyingPrice, exercisePrice, yearsToExpiry, interest, volatility, dividendYield))
    End Function

    Public Function PutRho(ByVal underlyingPrice As Double, ByVal exercisePrice As Double, ByVal daysToExpiry As Double, ByVal interest As Double, ByVal volatility As Double, ByVal dividendYield As Double) As Double
        If Not underlyingPrice > 0 Then Return 0
        If Not exercisePrice > 0 Then Return 0
        If Not daysToExpiry > 0 Then Return 0
        If Not interest > -1 Then Return 0
        If Not volatility > 0 Then Return 0

        'Convert Days to Years
        Dim yearsToExpiry As Double = daysToExpiry / 365

        'Convert Numbers to %
        interest = interest / 100
        volatility = volatility / 100
        dividendYield = dividendYield / 100

        Return -0.01 * exercisePrice * yearsToExpiry * Math.Exp(-interest * yearsToExpiry) * (1 - NormSDist(dTwo(underlyingPrice, exercisePrice, yearsToExpiry, interest, volatility, dividendYield)))
    End Function

    Public Function OptionGamma(ByVal underlyingPrice As Double, ByVal exercisePrice As Double, ByVal daysToExpiry As Double, ByVal interest As Double, ByVal volatility As Double, ByVal dividendYield As Double) As Double
        If Not underlyingPrice > 0 Then Return 0
        If Not exercisePrice > 0 Then Return 0
        If Not daysToExpiry > 0 Then Return 0
        If Not interest > -1 Then Return 0
        If Not volatility > 0 Then Return 0

        'Convert Days to Years
        Dim yearsToExpiry As Double = daysToExpiry / 365

        'Convert Numbers to %
        interest = interest / 100
        volatility = volatility / 100
        dividendYield = dividendYield / 100

        Return NdOne(underlyingPrice, exercisePrice, yearsToExpiry, interest, volatility, dividendYield) / (underlyingPrice * (volatility * Math.Sqrt(yearsToExpiry)))
    End Function

    Public Function OptionVega(ByVal underlyingPrice As Double, ByVal exercisePrice As Double, ByVal daysToExpiry As Double, ByVal interest As Double, ByVal volatility As Double, ByVal dividendYield As Double) As Double
        If Not underlyingPrice > 0 Then Return 0
        If Not exercisePrice > 0 Then Return 0
        If Not daysToExpiry > 0 Then Return 0
        If Not interest > -1 Then Return 0
        If Not volatility > 0 Then Return 0

        'Convert Days to Years
        Dim yearsToExpiry As Double = daysToExpiry / 365

        'Convert Numbers to %
        interest = interest / 100
        volatility = volatility / 100
        dividendYield = dividendYield / 100

        Return 0.01 * underlyingPrice * Math.Sqrt(yearsToExpiry) * NdOne(underlyingPrice, exercisePrice, yearsToExpiry, interest, volatility, dividendYield)
    End Function

    Private Function NormSDist(ByVal x As Double) As Double
        Dim t As Double
        Const b1 As Double = 0.31938153
        Const b2 As Double = -0.356563782
        Const b3 As Double = 1.781477937
        Const b4 As Double = -1.821255978
        Const b5 As Double = 1.330274429
        Const p As Double = 0.2316419
        Const c As Double = 0.39894228

        If x >= 0 Then
            t = 1.0 / (1.0 + p * x)
            Return (1.0 - c * Math.Exp(-x * x / 2.0) * t * (t * (t * (t * (t * b5 + b4) + b3) + b2) + b1))
        Else
            t = 1.0 / (1.0 - p * x)
            Return (c * Math.Exp(-x * x / 2.0) * t * (t * (t * (t * (t * b5 + b4) + b3) + b2) + b1))
        End If
    End Function

    Private Function dOne(ByVal underlyingPrice As Double, ByVal exercisePrice As Double, ByVal yearsToExpiry As Double, ByVal interest As Double, ByVal volatility As Double, ByVal dividendYield As Double) As Double
        Return (Math.Log(underlyingPrice / exercisePrice) + (interest - dividendYield + 0.5 * volatility ^ 2) * yearsToExpiry) / (volatility * (Math.Sqrt(yearsToExpiry)))
    End Function

    Private Function NdOne(ByVal underlyingPrice As Double, ByVal exercisePrice As Double, ByVal yearsToExpiry As Double, ByVal interest As Double, ByVal volatility As Double, ByVal dividendYield As Double) As Double
        Return Math.Exp(-(dOne(underlyingPrice, exercisePrice, yearsToExpiry, interest, volatility, dividendYield) ^ 2) / 2) / (Math.Sqrt(2 * 3.14159265358979))
    End Function

    Private Function dTwo(ByVal underlyingPrice As Double, ByVal exercisePrice As Double, ByVal yearsToExpiry As Double, ByVal interest As Double, ByVal volatility As Double, ByVal dividendYield As Double) As Double
        Return dOne(underlyingPrice, exercisePrice, yearsToExpiry, interest, volatility, dividendYield) - volatility * Math.Sqrt(yearsToExpiry)
    End Function

    Private Function NdTwo(ByVal underlyingPrice As Double, ByVal exercisePrice As Double, ByVal yearsToExpiry As Double, ByVal interest As Double, ByVal volatility As Double, ByVal dividendYield As Double) As Double
        Return NormSDist(dTwo(underlyingPrice, exercisePrice, yearsToExpiry, interest, volatility, dividendYield))
    End Function
End Class
