'----------------------------------------------------------------------
' Module: ErlangC Workforce Management Functions
' Purpose: Provides functions for calculating call center staffing requirements,
'          service levels, and related metrics using the Erlang C formula.
' Usage:   Call these functions from Excel or VBA to estimate required agents,
'          service levels, and other workforce management KPIs.
' Assumptions:
'   - Inputs are properly validated (e.g., non-negative values).
'   - Time units are consistent (e.g., seconds or minutes as specified).
'----------------------------------------------------------------------

Option Explicit

Function TrafficIntensity(Calls As Double, _
                          Reporting_Period As Long, _
                          AverageHandlingTime As Double) As Double

If Calls <= 0 Or Reporting_Period <= 0 Or AverageHandlingTime <= 0 Then
    TrafficIntensity = 0
Else
    TrafficIntensity = (Calls / (Reporting_Period * 60#)) * AverageHandlingTime
End If

End Function

Function AdjustedCallsForAbandonment(Calls As Long, AbandonmentRate As Double) As Double
    If AbandonmentRate < 0# Then AbandonmentRate = 0#
    If AbandonmentRate > 1# Then AbandonmentRate = 1#
    AdjustedCallsForAbandonment = Calls * (1 - AbandonmentRate)
End Function

' LogFact is a helper function to compute the logarithmic factorial using GammaLn approximation for mid-range n.
Function LogFact(N As Long) As Double
    If N < 2 Then
        LogFact = 0
    Else
        LogFact = WorksheetFunction.GammaLn(N + 1)
    End If

End Function


Function ErlangCTopRow(Intensity As Double, Agents As Long) As Double
    If Agents <= 170 Then
        ErlangCTopRow = (Intensity ^ Agents) / Application.WorksheetFunction.Fact(Agents)
    Else
        ' Use logarithmic factorial (Stirling's approximation) for large agents to avoid overflow
        ErlangCTopRow = Exp(Agents * Log(Intensity) - LogFact(Agents))
    End If
End Function

Function ErlangCBottomRow(Intensity As Double, Agents As Long) As Double
    Dim k As Long
    Dim answer As Double

    answer = 0
    For k = 0 To Agents - 1
        If k <= 170 Then
            answer = answer + ((Intensity ^ k) / Application.WorksheetFunction.Fact(k))
        Else
            answer = answer + Exp(k * Log(Intensity) - LogFact(k))
        End If
    Next k
    ErlangCBottomRow = answer

End Function

Function Occupancy(Intensity As Double, Agents As Long) As Double
    Dim result As Double

    If Agents <= 0 Then
        result = 0
    Else
        result = Intensity / Agents
    End If

    If result < 0 Then result = 0
    ' If result > 1, optionally handle over-occupancy here (e.g., raise an error or log a warning)
    ' Example: Debug.Print "Warning: Occupancy exceeds 1 (" & result & ")"
    ' Do not cap Occupancy at 1; return the actual value

    Occupancy = result

End Function

Function ErlangC(Intensity As Double, Agents As Long) As Double

Dim top As Double

    If Agents <= 0 Then
        ErlangC = 0
        Exit Function
    End If


    top = ErlangCTopRow(Intensity, Agents)
    If top = 0 Then
        ErlangC = 0
    Else
        ErlangC = top / (top + ((1 - Occupancy(Intensity, Agents)) * ErlangCBottomRow(Intensity, Agents)))
    End If

    If ErlangC < 0 Then ErlangC = 0
    If ErlangC > 1 Then ErlangC = 1

End Function

Function ProbabilityCallWaits(Intensity As Double, Agents As Long) As Double

Dim Occupancy As Double
Dim A As Double
Dim SumA As Double
Dim k As Long

If Agents <= 0 Or Intensity <= 0 Then
    ProbabilityCallWaits = 0
    Exit Function
End If

Occupancy = Intensity / Agents
A = 1
SumA = 0

For k = Agents To 1 Step -1
    A = A * k / Intensity
    SumA = SumA + A
Next k

ProbabilityCallWaits = 1 / (1 + ((1 - Occupancy) * SumA))

If ProbabilityCallWaits < 0 Then ProbabilityCallWaits = 0
If ProbabilityCallWaits > 1 Then ProbabilityCallWaits = 1

End Function


Function ServiceLevel(Calls As Long, _
                      Reporting_Period As Long, _
                      AverageHandlingTime As Double, _
                      ServiceLevelTime As Double, _
                      Agents As Long, _
                      Optional AbandonmentRate As Double = 0#) As Double

Dim AdjustedCalls As Double
Dim Intensity As Double

AdjustedCalls = AdjustedCallsForAbandonment(Calls, AbandonmentRate)
Intensity = TrafficIntensity(AdjustedCalls, Reporting_Period, AverageHandlingTime)

If Agents <= 0 Or Intensity <= 0 Then
    ServiceLevel = 0
    Exit Function
End If

ServiceLevel = 1 - _
    (ProbabilityCallWaits(Intensity, Agents) * _
     Exp(-(Agents - Intensity) * ServiceLevelTime / AverageHandlingTime))

If ServiceLevel < 0 Then ServiceLevel = 0
If ServiceLevel > 1 Then ServiceLevel = 1

End Function

Function FindMinAgents(Calls As Long, _
                       Reporting_Period As Long, _
                       AverageHandlingTime As Double, _
                       ServiceLevelTarget As Double, _
                       ServiceLevelTime As Double, _
                       Optional AbandonmentRate As Double = 0#) As Long

Dim AdjustedCalls As Double
Dim Intensity As Double
Dim Low As Long, High As Long, Mid As Long

AdjustedCalls = AdjustedCallsForAbandonment(Calls, AbandonmentRate)
Intensity = TrafficIntensity(AdjustedCalls, Reporting_Period, AverageHandlingTime)

If Intensity <= 0 Then
    FindMinAgents = 0
    Exit Function
End If

Low = Application.WorksheetFunction.Max(1, Int(Intensity))
High = Low

Do While ServiceLevel(Calls, Reporting_Period, AverageHandlingTime, ServiceLevelTime, High, AbandonmentRate) < ServiceLevelTarget
    High = High * 2
    If High > 10000 Then Exit Do
Loop

Do While Low < High
    Mid = (Low + High) \\ 2
    If ServiceLevel(Calls, Reporting_Period, AverageHandlingTime, ServiceLevelTime, Mid, AbandonmentRate) < ServiceLevelTarget Then
        Low = Mid + 1
    Else
        High = Mid
    End If
Loop

FindMinAgents = Low

End Function

Function AgentsRequired(Calls As Long, _
                        Reporting_Period As Long, _
                        AverageHandlingTime As Double, _
                        ServiceLevelTarget As Double, _
                        ServiceLevelTime As Double, _
                        Optional AbandonmentRate As Double = 0#) As Double

Dim Intensity As Double
Dim N As Long
Dim SL_Low As Double, SL_High As Double
Dim Fraction As Double

Intensity = TrafficIntensity(AdjustedCallsForAbandonment(Calls, AbandonmentRate), Reporting_Period, AverageHandlingTime)

If Intensity = 0 Then
    AgentsRequired = 0
    Exit Function
End If

If Intensity < 1 Then
    AgentsRequired = Intensity
    Exit Function
End If

N = FindMinAgents(Calls, Reporting_Period, AverageHandlingTime, _
                  ServiceLevelTarget, ServiceLevelTime, AbandonmentRate)

SL_High = ServiceLevel(Calls, Reporting_Period, AverageHandlingTime, ServiceLevelTime, N, AbandonmentRate)
SL_Low = ServiceLevel(Calls, Reporting_Period, AverageHandlingTime, ServiceLevelTime, N - 1, AbandonmentRate)

If SL_High = SL_Low Then
    AgentsRequired = N
Else
    Fraction = (ServiceLevelTarget - SL_Low) / (SL_High - SL_Low)
    If Fraction < 0 Then Fraction = 0
    If Fraction > 1 Then Fraction = 1
    AgentsRequired = (N - 1) + Fraction
End If

End Function

' Function to calculate required FTE agents considering shrinkage
' shrinkage: Fraction of time agents are unavailable (e.g., 0.2 for 20%)
' abandonment rate: Fraction of calls that abandon before service (e.g., 0.2 for 20%)
Function AgentsFTE(Calls As Long, _
                           Reporting_Period As Long, _
                           AverageHandlingTime As Double, _
                           ServiceLevelTarget As Double, _
                           ServiceLevelTime As Double, _
                           Shrinkage As Double, _
                           Optional AbandonmentRate As Double = 0#) As Double

Dim BaseAgents As Double

If Shrinkage < 0 Then Shrinkage = 0
If Shrinkage >= 1 Then Shrinkage = 0.99

BaseAgents = AgentsRequired(Calls, Reporting_Period, _
                            AverageHandlingTime, _
                            ServiceLevelTarget, _
                            ServiceLevelTime, _
                            AbandonmentRate)

If BaseAgents = 0 Then
    AgentsFTE = 0
Else
    AgentsFTE = BaseAgents / (1 - Shrinkage)
End If

End Function
'----------------------------------------------------------------------
' End of ErlangC Workforce Management Functions Module
' This module provides Erlang C-based calculations for call center staffing,
' service levels, and related workforce management metrics.
' Functions are intended for use in Excel or VBA environments.
'----------------------------------------------------------------------
