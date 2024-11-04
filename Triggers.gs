/**
 * Verifica se existe algum acionador da função fornecida. Caso exista a função retorna 1, caso não retorna 0.
 * 
 * @param {function} name_Func - Entre com o nome da função que deseja verificar se existe o acionador.
 */
function Triggers_Match(name_Func){
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
function Triggers_Delete(name_Func){
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
function Trigger_Create(name_Func){
  ScriptApp.newTrigger(name_Func)
    .timeBased()
    .everyMinutes(1)
    .create();
}
