Attribute VB_Name = "Lamerizer"
Option Explicit
Public Function Lamerized(Text As String) As String
Text = Replace(Text, "a", "�")
Text = Replace(Text, "b", "�")
Text = Replace(Text, "c", "�")
Text = Replace(Text, "d", "�")
Text = Replace(Text, "e", "�")
Text = Replace(Text, "f", "F")
Text = Replace(Text, "g", "G")
Text = Replace(Text, "h", "H")
Text = Replace(Text, "i", "�")
Text = Replace(Text, "j", "J")
Text = Replace(Text, "k", "K")
Text = Replace(Text, "l", "�")
Text = Replace(Text, "m", "M")
Text = Replace(Text, "n", "�")
Text = Replace(Text, "o", "�")
Text = Replace(Text, "p", "�")
Text = Replace(Text, "q", "Q")
Text = Replace(Text, "r", "R")
Text = Replace(Text, "s", "�")
Text = Replace(Text, "t", "T")
Text = Replace(Text, "u", "�")
Text = Replace(Text, "v", "V")
Text = Replace(Text, "w", "W")
Text = Replace(Text, "x", "�")
Text = Replace(Text, "y", "�")
Text = Replace(Text, "z", "Z")
Lamerized = Text
End Function


