*** Settings ***
Documentation     Proyecto que busca automatizar consultas hacia Civil Pjud.
...               Se necesitan las librerias instaladas para poder hacer funcionar el script.
Library           SeleniumLibrary
Library           ExcelLibrary
Library           clipboard
Library           String
Library           DateTime
Library           Collections

*** Variables ***
${Url}            https://laboral.pjud.cl/    # Direccion de la pagina a realizar las consultas
${PathExcel}      resultado/Nombres.xls    #Ubicacion de archivo Excel a consultar.
${NombreHojaExcel}    Laboral    #Nombre de la hoja excel que se consulta.
${ContadorTribunalOrigen}    0
@{TotalJuzgado1}    1348    6    1349
@{TotalJuzgado}    6    13    14    26    27    29    34    36    37    46    47    48    49    50    51    52    53
...               83    84    85    86    87    88    89    90    94    96    97    98    99    101    102    103    111
...               113    114    115    116    117    119    126    127    132    133    135    136    138    139    140    141    147
...               149    150    151    152    157    158    159    160    187    188    189    190    191    192    193    194    195
...               196    204    206    207    208    209    210    211    212    213    214    215    216    222    223    224    225
...               226    227    238    240    241    243    244    245    248    249    250    257    258    373    374    375    377
...               378    385    386    387    388    659    660    946    947    996    1013    1150    1151    1152    1333    1334    1335
...               1336    1337    1338    1339    1340    1341    1342    1343    1344    1345    1346    1347    1348    1349    1351    1352    1357
...               1358    1359    1360    1361    1362    1363    1500    1501    1502
${Contador}       1    #Contador que recorrera el total de valores de archivo excel.
${NombreCopiar}    ${EMPTY}    #Nombre que se extrae de excel
${ApellidoPaternoCopiar}    ${EMPTY}
${ApellidoMaternoCopiar}    ${EMPTY}
${RutCopiar}      ${EMPTY}
${ContadorCasos}    ${EMPTY}
${ContadorDeCasosInternos}    1
${SiTienenCaso}    ${EMPTY}
${CounterInside}    1
${Var12}          1
@{ParaGuardarEnExcelSumarizado}    # Listado Para guardar En excel
${ParaGuardarEnExcelSumarizado}    ${EMPTY}

*** Test Cases ***
TestFinal
    Open Excel    ${PathExcel}
    ${Count1}    Get Row Count    ${NombreHojaExcel}    #Total de filas
    @{Count1}    Get column values    ${NombreHojaExcel}    1    #Valores de la columna 1
    FOR    ${Var1}    IN    @{Count1}    #Recorre    cada fila de archivo excel
        BuscadorDeCasos
        log    ${Contador}
        Sleep    5s
        AumentadorDeNumeroPorCaso
        Log    ${Var1}
    END
    Log List    ${ParaGuardarEnExcelSumarizado}

*** Keywords ***
BuscadorDeCasos
    [Documentation]    Rescata variables desde Excel.
    Open Excel    ${PathExcel}
    Open Browser    ${Url}    chrome    #Apertura de explorador
    Sleep    10s    \    #Espera de 10 segundos
    Select Frame    name=body
    Click Element    //td[contains(@id,'tdCuatro')]
    log    ${Contador}
    ${NombreCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    B${Contador}
    clipboard.Copy    ${NombreCopiar}
    ${NombreCopiar}    Set Suite Variable    ${NombreCopiar}
    Log    ${NombreCopiar}
    Click Element    //input[contains(@name,'NOM_Consulta')]
    Press Keys    none    CTRL+V
    ${ApellidoPaternoCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    D${Contador}
    clipboard.Copy    ${ApellidoPaternoCopiar}
    ${ApellidoPaternoCopiar}    Set Suite Variable    ${ApellidoPaternoCopiar}
    Log    ${ApellidoPaternoCopiar}
    Click Element    //input[contains(@name,'APE_Paterno')]
    Press Keys    none    CTRL+V
    ${ApellidoMaternoCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    E${Contador}
    clipboard.Copy    ${ApellidoMaternoCopiar}
    ${ApellidoMaternoCopiar}    Set Suite Variable    ${ApellidoMaternoCopiar}
    Log    ${ApellidoMaternoCopiar}
    Click Element    //input[contains(@name,'APE_Materno')]
    Press Keys    none    CTRL+V
    ${RutCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    A${Contador}
    clipboard.Copy    ${RutCopiar}
    ${RutCopiar}    Set Suite Variable    ${RutCopiar}
    Log    ${RutCopiar}
    Sleep    10s
    Click Element    //img[@onclick='document.AtPublicoPpalForm.irAccionAtPublico.click();']
    Sleep    18s
    FOR    ${VAR}    IN    @{TotalJuzgado1}
        Sleep    15s
        Select From List By Value    name:COD_TribunalSinTodos    ${VAR}
        Sleep    5s
        log    ${VAR}
        Click Element    //img[@onclick='document.AtPublicoPpalForm.irAccionAtPublico.click();']
        Sleep    10s
        ${ContadorCasoExiste}=    Get Element Count    (//a[@onclick='ValDobleSubmit()'])[1]
        Sleep    10s
        Run Keyword If    ${ContadorCasoExiste}>0    RecorrerCasosInternos
        ...    ELSE    log    "No existe Registro Valido"
        Sleep    10s
    END
    Close Browser

AumentadorDeNumeroPorCaso
    [Documentation]    Contador del total de personas de los cuales se consideraran para las consultas.
    ${temp}    Evaluate    ${Contador} + 1
    Set Test Variable    ${Contador}    ${temp}

GuardadorEnExcel
    [Documentation]    Generador de archivo Excel con fecha y hora respectiva
    Open Excel    resultado/Prototipo.xls
    Sleep    5s
    log    ${Contador}
    log    ${RutCopiar}
    log    ${NombreCopiar}
    log    ${ApellidoPaternoCopiar}
    log    ${ApellidoMaternoCopiar}
    Put String To Cell    resultado    0    ${Contador}    ${RutCopiar}
    Put String To Cell    resultado    1    ${Contador}    ${NombreCopiar}
    Put String To Cell    resultado    2    ${Contador}    ${ApellidoPaternoCopiar}
    Put String To Cell    resultado    3    ${Contador}    ${ApellidoMaternoCopiar}
    ${SiTienenCaso}    Get WebElements    (//td[contains(@height,'11')])[1]
    ${NombreJuzgado}    Get WebElements    //td[contains(@width,'392')]
    ${NumeroCaso}=    Get Text    ${SiTienenCaso[0]}
    ${NombreJuzgado1}=    Get Text    ${NombreJuzgado[0]}
    Put String To Cell    resultado    4    ${Contador}    ${NumeroCaso}
    Put String To Cell    resultado    5    ${Contador}    ${NombreJuzgado1}
    ${timestamp} =    Get Current Date    result_format=%Y-%m-%d-%H-%M
    ${filename} =    Set Variable    resultado-${timestamp}.xls
    Save Excel    resultado/${filename}
    Append To List    ${ParaGuardarEnExcelSumarizado}    ${RutCopiar}
    Append To List    ${ParaGuardarEnExcelSumarizado}    ${NombreCopiar}
    Append To List    ${ParaGuardarEnExcelSumarizado}    ${ApellidoPaternoCopiar}
    Append To List    ${ParaGuardarEnExcelSumarizado}    ${ApellidoMaternoCopiar}
    Append To List    ${ParaGuardarEnExcelSumarizado}    ${NumeroCaso}
    Append To List    ${ParaGuardarEnExcelSumarizado}    ${NombreJuzgado1}

ValidarRutExcelHaciaPjud
    [Documentation]    Se valida que exista un caso valido para el rut mostrado.
    log    ${ContadorCasos}
    log    ${NombreCopiar}
    log    ${RutCopiar}
    log    ${ApellidoMaternoCopiar}
    log    ${ApellidoPaternoCopiar}
    Sleep    7s
    Sleep    2s
    ${Span}=    SeleniumLibrary.Get WebElements    (//td[contains(.,'${RutCopiar}')])[3]
    Log    ${Span}
    log    ${RutCopiar}
    ${test}=    Get Element Count    (//td[contains(.,'${RutCopiar}')])[3]
    log    ${test}
    Sleep    8s
    Run Keyword If    ${test}>0    GuardadorEnExcel
    ...    ELSE    log    "No existe Registro Valido"

RecorrerCasosInternos
    Sleep    10s
    FOR    ${Var12}    IN RANGE    9999
        log    ${CounterInside}
        ${Var13}=    Get Element Count    (//a[@onclick='ValDobleSubmit()'])[${CounterInside}]
        log    ${Var13}
        Convert To Number    ${Var13}
        Exit For Loop If    ${Var13}==0
        Click Element    (//a[@onclick='ValDobleSubmit()'])[${CounterInside}]
        Sleep    10s
        ${CounterInside}=    Evaluate    ${CounterInside}+1
        Click Element    (//td[contains(.,'Litigantes')])[1]
        Sleep    5s
        ValidarRutExcelHaciaPjud
        go back
        Select Frame    name=body
        Sleep    10s
    END
