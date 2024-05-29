function doGet() {
  plantilla = HtmlService.createTemplateFromFile('index');
  //return plantilla.evaluate().setTitle("Buscador de Horarios") 
  var html = HtmlService.createTemplateFromFile("index");
  var evaluated = html.evaluate();
  evaluated.setTitle("Buscador de Horarios");
  evaluated.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return evaluated.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)

}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}

var agentes = []
var fechasBuscador = []
var scores = []
var matrixId
function abrirMatriz(ag,mes) {
  let agentes = queryAgentes()
  
  if(mes=="4"){
    matrixId="xx"
  }
  if(mes=="3"){
    matrixId="xx"
  }
  if(mes=="5"){
    matrixId="xx"
  }

  let score = ""
  for (i = 1; i < agentes.length; i++) {
    if (agentes[i][0] == ag) {
      score = agentes[i][1]
      if (score == "CNYoizen") {
        score = "Corredor Norte"
      }
      if (score == "COMBO") {
        score = "COMBO-CM"
      }
    }
  }
  if (score != "") {



    let hoja = SpreadsheetApp.openById(matrixId).getSheetByName(score).getDataRange().getDisplayValues()
    let hojas = []

    let SS = SpreadsheetApp.getActive()
    let ws = SS.getSheetByName("registro");
    let date = new Date();

    ws.appendRow([
      ag,
      score,
      date
    ]);

    let headers = []
    let datos = []
    let horarios = new Object()
    hoja.forEach(fila => {
      for (i = 1; i < fila.length; i++) {
        if (fila[i] == ag) {
          for (j = 17; j < fila.length; j++) {
            if (fila[j] !== "CN") {

              if ((horarios[hoja[1][j]]) != "") {
                if (fila[j] == "F") {
                  horarios[hoja[1][j]] = `Franco`
                } else {
                  if (fila[j] == "V") {
                    horarios[hoja[1][j]] = `Vacaciones`
                  } else {
                    if (fila[j] == "NP") {
                      horarios[hoja[1][j]] = `No Program.`
                    } else {
                      if (fila[j] == "Fe") {
                        horarios[hoja[1][j]] = `Feriado`
                      } else {
                        horarios[hoja[1][j]] = `De ${fila[j]} a ${fila[j + 1]}`
                        //Logger.log(hoja[1][j])
                        let date2 = new Date(hoja[1][j])
                        Logger.log(date2)
                      }
                    }

                  }
                }
              }

            }


          }

        }
      }



    })
    var respuesta
    // Logger.log(horarios)
    function showProps(obj, objName) {
      var result = [];
      for (var i in obj) {
        // obj.hasOwnProperty() se usa para filtrar propiedades de la cadena de prototipos del objeto
        if (obj.hasOwnProperty(i)) {
          if (i != "") {
            if (obj[i] == "De F a F") {
              result.push(`${i}  Franco`);
            }
            result.push(`${i}  ${obj[i]}`);

          }
        }
      }

      return result;
    }

    respuesta = showProps(horarios, "horarios")
    //Logger.log(respuesta)
    return respuesta
  } else {
    return "NP"
  }

}

//Funcion para traer nombres de agentes de la matriz
function queryAgentes() {
  var listaAgentes = [];
  let formula = `=QUERY({IMPORTRANGE(Matriz,CATV);IMPORTRANGE(Matriz,COMBO);IMPORTRANGE(Matriz,SCORE1);IMPORTRANGE(Matriz,FACTURA);IMPORTRANGE(Matriz,CORREDOR_NORTE)},"SELECT Col4,Col10 WHERE Col4 != 'NOMBRE' AND Col3 != '' AND Col4 != 'AGENTE'",0)`
  let libro = SpreadsheetApp.getActiveSpreadsheet()
  let hojaAgentes = libro.getSheetByName("ListaAgentes")
  hojaAgentes.getRange("A1").setFormula(formula);
  let agentes = hojaAgentes.getDataRange().getValues()
  agentes.forEach(agente => {
    listaAgentes.push(agente)
  })
  //Logger.log(listaAgentes)
  return listaAgentes

}

function horarioXFecha(dato) {
  let fecha;

  var today = new Date(dato);
  today.setDate(today.getDate() + 1);
  if (today.getDate() < 10) {
    fecha = (today.toLocaleDateString("es-AR", { weekday: 'short' }) + " 0" + today.toLocaleDateString("es-AR", { day: 'numeric' }) + "-" + today.toLocaleDateString("es-AR", { month: 'short' }));
  } else {
    fecha = (today.toLocaleDateString("es-AR", { weekday: 'short' }) + " " + today.toLocaleDateString("es-AR", { day: 'numeric' }) + "-" + today.toLocaleDateString("es-AR", { month: 'short' }));
  }
  //Logger.log(fecha)

  let hojaCatv = SpreadsheetApp.openById("xxxx").getSheetByName("CATV").getDataRange().getDisplayValues()
  let hojaFactura = SpreadsheetApp.openById("xxxx").getSheetByName("FACTURA").getDataRange().getDisplayValues()
  let hojaCN = SpreadsheetApp.openById("xxxx").getSheetByName("CORREDOR NORTE").getDataRange().getDisplayValues()
  let hojaCombo = SpreadsheetApp.openById("xxxx").getSheetByName("COMBO-CM").getDataRange().getDisplayValues()
  let hojaScore1 = SpreadsheetApp.openById("xxxx").getSheetByName("SCORE1").getDataRange().getDisplayValues()

  let horarios = []
  let indexFecha = 0

  Logger.log(fecha)

  hojaFactura.forEach(fila => {

    for (i = 0; i < fila.length; i++) {
      //Logger.log(fila[i])
      if (fila[i] == fecha) {
        indexFecha = i
        return
      }
    }

  })
  Logger.log(indexFecha)
  hojaCatv.forEach(fila => {

    if (fila[3] !== "") {
      if (fila[3] !== "AGENTE") {
        if (fila[3] !== "NOMBRE") {
          let agente = {
            nombre: fila[3],
            sup: fila[4],
            score: fila[7],
            horarioIn: fila[indexFecha],
            horarioOut: fila[indexFecha + 1]
          }
          horarios.push(agente)

        }
      }

    }


  })

  hojaCN.forEach(fila => {

    if (fila[3] !== "") {
      if (fila[3] !== "AGENTE") {
        if (fila[3] !== "NOMBRE") {
          let agente = {
            nombre: fila[3],
            sup: fila[4],
            score: fila[7],
            horarioIn: fila[indexFecha],
            horarioOut: fila[indexFecha + 1]
          }
          horarios.push(agente)

        }
      }

    }




  })
  hojaCombo.forEach(fila => {

    if (fila[3] !== "") {
      if (fila[3] !== "AGENTE") {
        if (fila[3] !== "NOMBRE") {
          let agente = {
            nombre: fila[3],
            sup: fila[4],
            score: fila[7],
            horarioIn: fila[indexFecha],
            horarioOut: fila[indexFecha + 1]
          }
          horarios.push(agente)

        }
      }

    }




  })
  hojaFactura.forEach(fila => {

    if (fila[3] !== "") {
      if (fila[3] !== "AGENTE") {
        if (fila[3] !== "NOMBRE") {
          let agente = {
            nombre: fila[3],
            sup: fila[4],
            score: fila[7],
            horarioIn: fila[indexFecha],
            horarioOut: fila[indexFecha + 1]
          }
          horarios.push(agente)

        }
      }

    }



  })
  hojaScore1.forEach(fila => {

    if (fila[3] !== "") {
      if (fila[3] !== "AGENTE") {
        if (fila[3] !== "NOMBRE") {
          let agente = {
            nombre: fila[3],
            sup: fila[4],
            score: fila[7],
            horarioIn: fila[indexFecha],
            horarioOut: fila[indexFecha + 1]
          }
          horarios.push(agente)

        }
      }

    }





  })

  return horarios


}

function horarioAgente(agente = "s/d") {
  let fecha;
  let agenteEncontrado = false

  var today = new Date("03-05-2024");
  today.setDate(today.getDate() + 1);
  Logger.log(today)
  if (today.getDate() < 10) {
    fecha = (today.toLocaleDateString("es-AR", { weekday: 'short' }) + " 0" + today.toLocaleDateString("es-AR", { day: 'numeric' }) + "-" + today.toLocaleDateString("es-AR", { month: 'short' }));
  } else {
    fecha = (today.toLocaleDateString("es-AR", { weekday: 'short' }) + " " + today.toLocaleDateString("es-AR", { day: 'numeric' }) + "-" + today.toLocaleDateString("es-AR", { month: 'short' }));
  }
  Logger.log(fecha)

  let hojaCatv = SpreadsheetApp.openById("xxx").getSheetByName("CATV").getDataRange().getDisplayValues()
  let hojaFactura = SpreadsheetApp.openById("xxx").getSheetByName("FACTURA").getDataRange().getDisplayValues()
  let hojaCN = SpreadsheetApp.openById("xxx").getSheetByName("CORREDOR NORTE").getDataRange().getDisplayValues()
  let hojaCombo = SpreadsheetApp.openById("xxx").getSheetByName("COMBO-CM").getDataRange().getDisplayValues()
  let hojaScore1 = SpreadsheetApp.openById("xxx").getSheetByName("SCORE1").getDataRange().getDisplayValues()

  let horarios = []
  let indexFecha = 0

  Logger.log(fecha)

  hojaFactura.forEach(fila => {

    for (i = 0; i < fila.length; i++) {
      //Logger.log(fila[i])
      if (fila[i] == fecha) {
        indexFecha = i
        return
      }
    }

  })
  Logger.log(indexFecha)
  hojaCatv.forEach(fila => {

    if (fila[3] !== "") {
      if (fila[3] !== "AGENTE") {
        if (fila[3] !== "NOMBRE") {
          let agente = {
            nombre: fila[3],
            sup: fila[4],
            score: fila[7],
            horarioIn: fila[indexFecha],
            horarioOut: fila[indexFecha + 1]
          }
          horarios.push(agente)

        }
      }

    }


  })

  hojaCN.forEach(fila => {

    if (fila[3] !== "") {
      if (fila[3] !== "AGENTE") {
        if (fila[3] !== "NOMBRE") {
          let agente = {
            nombre: fila[3],
            sup: fila[4],
            score: fila[7],
            horarioIn: fila[indexFecha],
            horarioOut: fila[indexFecha + 1]
          }
          horarios.push(agente)

        }
      }

    }




  })
  hojaCombo.forEach(fila => {

    if (fila[3] !== "") {
      if (fila[3] !== "AGENTE") {
        if (fila[3] !== "NOMBRE") {
          let agente = {
            nombre: fila[3],
            sup: fila[4],
            score: fila[7],
            horarioIn: fila[indexFecha],
            horarioOut: fila[indexFecha + 1]
          }
          horarios.push(agente)

        }
      }

    }




  })
  hojaFactura.forEach(fila => {

    if (fila[3] !== "") {
      if (fila[3] !== "AGENTE") {
        if (fila[3] !== "NOMBRE") {
          let agente = {
            nombre: fila[3],
            sup: fila[4],
            score: fila[7],
            horarioIn: fila[indexFecha],
            horarioOut: fila[indexFecha + 1]
          }
          horarios.push(agente)

        }
      }

    }



  })
  hojaScore1.forEach(fila => {

    if (fila[3] !== "") {
      if (fila[3] !== "AGENTE") {
        if (fila[3] !== "NOMBRE") {
          let agente = {
            nombre: fila[3],
            sup: fila[4],
            score: fila[7],
            horarioIn: fila[indexFecha],
            horarioOut: fila[indexFecha + 1]
          }
          horarios.push(agente)

        }
      }

    }





  })

  return horarios


}

function buscarCambiosFranco(ag = "Bergagno, Oriana Berenisse") {
  let agentes = queryAgentes()
  let score = ""
  for (i = 1; i < agentes.length; i++) {
    if (agentes[i][0] == ag) {
      score = agentes[i][1]
      if (score == "CNYoizen") {
        score = "Corredor Norte"
      }
    }
  }
  if (score != "") {
    let hoja = SpreadsheetApp.openById("xxx").getSheetByName(score).getDataRange().getDisplayValues()
    let indexFila
    hoja.forEach(fila => {
      //let date = hoja[i][j].slice(4,9)
      let auxFila = 0
      fila.forEach(celda => {

        if (celda.slice(4,10) == "24-abr") {
          indexFila = auxFila

        }
        auxFila++
      })

      if(fila[indexFila]=="F"){
        Logger.log(fila[3])

        Logger.log("-----------")

      }





    })




  }
}




