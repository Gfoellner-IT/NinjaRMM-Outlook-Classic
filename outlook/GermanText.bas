Attribute VB_Name = "GermanText"
' GermanText v0.1.0
'
' Source files are ASCII-only, but texts for the (German) user need umlauts. De() turns the markers
' [ae] [oe] [ue] [Ae] [Oe] [Ue] [ss] in a text into the real characters:
'   De("Pr[ue]fung")  gives "Pr", u umlaut, "fung"
' Pure function, no dependencies.

Option Explicit

Public Function De(ByVal text As String) As String
    text = Replace(text, "[ae]", ChrW$(228))
    text = Replace(text, "[oe]", ChrW$(246))
    text = Replace(text, "[ue]", ChrW$(252))
    text = Replace(text, "[Ae]", ChrW$(196))
    text = Replace(text, "[Oe]", ChrW$(214))
    text = Replace(text, "[Ue]", ChrW$(220))
    text = Replace(text, "[ss]", ChrW$(223))
    De = text
End Function
