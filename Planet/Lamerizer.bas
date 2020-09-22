Attribute VB_Name = "Lamerizer"
Option Explicit
Public Function Lamerized(Text As String) As String
Text = Replace(Text, "a", "â")
Text = Replace(Text, "b", "ß")
Text = Replace(Text, "c", "Ç")
Text = Replace(Text, "d", "Ð")
Text = Replace(Text, "e", "ë")
Text = Replace(Text, "f", "F")
Text = Replace(Text, "g", "G")
Text = Replace(Text, "h", "H")
Text = Replace(Text, "i", "í")
Text = Replace(Text, "j", "J")
Text = Replace(Text, "k", "K")
Text = Replace(Text, "l", "£")
Text = Replace(Text, "m", "M")
Text = Replace(Text, "n", "Ñ")
Text = Replace(Text, "o", "Ø")
Text = Replace(Text, "p", "þ")
Text = Replace(Text, "q", "Q")
Text = Replace(Text, "r", "R")
Text = Replace(Text, "s", "§")
Text = Replace(Text, "t", "T")
Text = Replace(Text, "u", "ú")
Text = Replace(Text, "v", "V")
Text = Replace(Text, "w", "W")
Text = Replace(Text, "x", "×")
Text = Replace(Text, "y", "¥")
Text = Replace(Text, "z", "Z")
Lamerized = Text
End Function


