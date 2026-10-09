"""Exécute les modules VBA dans LibreOffice (mode compatibilité VBA) pour détecter les erreurs.

Usage : python3 tester_libreoffice.py <classeur.xlsx> <dossier_vba> <sortie.xlsx> <journal.txt>
Les MsgBox sont redirigées vers le journal (pas d'interface graphique).
"""
import glob
import os
import re
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue

classeur, dossier_vba, sortie, journal = map(os.path.abspath, sys.argv[1:5])
macro = sys.argv[5] if len(sys.argv) > 5 else "LancerAnalyse"
# Dossier simulant Euronext : fichiers <ISIN>.csv renvoyés à la place du téléchargement
dossier_euronext = os.environ.get("TEST_EURONEXT_DIR", "")


def prop(nom, valeur):
    p = PropertyValue()
    p.Name, p.Value = nom, valeur
    return p


proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore",
                         "--accept=socket,host=localhost,port=2002;urp;"],
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
try:
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(60):
        try:
            ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
            break
        except Exception:
            time.sleep(0.5)
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(classeur), "_blank", 0,
                                       (prop("Hidden", True), prop("MacroExecutionMode", 4)))
    libs = doc.BasicLibraries
    if not libs.hasByName("Standard"):
        libs.createLibrary("Standard")
    lib = libs.getByName("Standard")
    journal_basic = journal.replace("\\", "/")
    for chemin in sorted(glob.glob(os.path.join(dossier_vba, "*.bas"))):
        nom = os.path.splitext(os.path.basename(chemin))[0]
        code = open(chemin, encoding="ascii").read()
        code = re.sub(r"^Attribute VB_Name.*\n", "", code, flags=re.M)
        code = re.sub(r"\bMsgBox (?=[^(=])", "TestLog ", code)
        code = re.sub(r"#If Mac Then.*?#End If\n", "", code, flags=re.S)  # spécifique Excel Mac
        if dossier_euronext:
            code = code.replace("Public Function TelechargerTexte(", "Public Function TelechargerTexte_Reel(")
        code = "Option VBASupport 1\n" + code
        if nom in ("modMain", "Module1_Lancement"):
            code += f'''
Public Sub TestLog(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant)
    Dim f As Integer
    f = FreeFile
    Open "{journal_basic}" For Append As #f
    Print #f, "MSG: " & a
    Close #f
End Sub
'''
        if dossier_euronext and nom in ("modMain", "Module1_Lancement"):
            code += f'''
Public Function TelechargerTexte(ByVal adresse As String) As String
    Dim isin As String, f As Integer, chemin As String
    TestLog "URL: " & adresse
    isin = Mid$(adresse, InStr(adresse, "getFullDownloadAjax/") + 20, 12)
    chemin = "{dossier_euronext}/" & isin & ".csv"
    If Dir(chemin) = "" Then
        TelechargerTexte = "<html>Page introuvable</html>"
        Exit Function
    End If
    f = FreeFile
    Open chemin For Input As #f
    TelechargerTexte = Input$(LOF(f), #f)
    Close #f
End Function
'''
        if lib.hasByName(nom):
            lib.removeByName(nom)
        lib.insertByName(nom, code)
    sp = doc.getScriptProvider()
    module_principal = os.environ.get("TEST_MODULE", "modMain")
    script = sp.getScript(f"vnd.sun.star.script:Standard.{module_principal}.{macro}?language=Basic&location=document")
    t0 = time.time()
    try:
        script.invoke((), (), ())
        print(f"Macro {macro} exécutée en {time.time() - t0:.1f} s")
    except Exception as e:
        print("ERREUR d'exécution :", e)
    doc.calculateAll()
    doc.storeToURL(uno.systemPathToFileUrl(sortie), (prop("FilterName", "Calc MS Excel 2007 XML"),))
    doc.close(True)
finally:
    proc.terminate()
