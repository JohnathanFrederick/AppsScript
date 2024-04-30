~~~javascript
/**
 * Verifica se existe algum acionador da função fornecida. Caso exista a função retorna 1, caso não retorna 0.
 * 
 * @param {function} name_Func - Entre com o nome da função que deseja verificar se existe o acionador.
 */
function Util_VerifiqueTriggers(name_Func){
  var triggers = ScriptApp.getProjectTriggers();
  for(let i = 0; i<triggers.length;i++){
    if(triggers[i].getHandlerFunction() === name_Func){
      return 1;
    }
  }
  return 0
}
/**
 * Deleta todos os acionadores de uma determinada função, passada como argumento à esta função.
 * 
 * @param {function} name_Func - Entre com o nome da função que deseje deletar os acionadores.
 */
function Util_DeleteTriggers(name_Func){
  var triggers = ScriptApp.getProjectTriggers();
  for(var i=0; i<triggers.length; i++){
    if(triggers[i].getHandlerFunction() === name_Func){
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
/**
 * Cria um acionador para uma função, baseado no tempo com a taxa de execução a cada minuto.
 * 
 * @param {function} name_Func - Entre com o nome da função que deseja criar o acionador
 */
function Util_CreateTriggers(name_Func){
  ScriptApp.newTrigger(name_Func)
    .timeBased()
    .everyMinutes(1)
    .create();
}
/**
 * Obtém os intervalos nomeados de uma spreadsheet e os retorna como um dicionário, em que as chaves do dicionário são os nomes dos intervalos nomeados e as propriedades de cada chave são um objeto de intervalo.
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss - A spreadsheet a obter os intervalos nomeados
 * 
 * @returns {SpreadsheetApp.Range} - Objeto range
 */
function Util_GetNamedRanges(ss){
  let array_NamedRanges = ss.getNamedRanges();
  let dic_Ranges = {};
  for(let obj_Range of array_NamedRanges){
    let name = obj_Range.getName();
    let range = obj_Range.getRange();
    dic_Ranges[name] = range;
  }
  return dic_Ranges;
}
~~~
