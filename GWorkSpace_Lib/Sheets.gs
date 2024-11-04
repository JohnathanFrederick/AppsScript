/**
 * A função executa uma função, passando como argumento um dicionário resultante de uma string JSON presente em cada célula de um determinado intervalo.
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss - A spreadsheet a ser processada. Este parâmetro é usado para mostrar uma caixa de mensagem com o método .toas()
 * @param {SpreadsheetApp.Range} range_DadosJson - Um inervalo coluna, em que cada célula contenha dados no formato JSON, a ser passado como argumento para uma função;
 * @param {SpreadsheetApp.Range} range_Status - Um intervalo coluna, em que cada célula contém o status de execução da linha. Este deve possuir o mesmo tamanho que o intervalo que contém os dados;
 * @param {SpreadsheetApp.Range} range_Erro - Um intervalo coluna, em que cada célula uma mensagem de erro com seus detalhes, caso o erro ocorra;
 * @param {SpreadsheetApp.Range} range_Info - Um intervalo coluna, em que cada célula o código colocará informações sobre a requisição;
 * @param {Function} func_Apply - A função que receberá os dados de cada célula como argumento;
 * @param {Function} func_Trigger - A função cujo acionadores serão deletados no final de todas as execuções;
 * @param {Function} func_Finally - A função que será executada quando todas as execuções foram concluídas. Caso este argumento seja null, esta ação será ignorada;
 * @param {Number} num_MaxExecutions - O número máximo que a função poderá ser executada. Por padrão este valor é igual a 1;
 * @param {Number} num_MaxTentativas - O número máximo que a função poderá tentar uma execução. Por padrão este valor é igual a 1;
 * @param {Number} time_MaxWaitLock - O tempo máximo (em milissegundos) a esperar para a liberação do semáforo. Por padrão este valor é igual a 30000.
 * @param {Number} depurar - Nível de depuração para a função. Por padrão este argumento é 0. Quando diferente de zero, a função imprimi os log's para depuração, tendo os seguintes níveis de depuração:
 *    - 1: Mensagens principais
 *    - 2: Valores principais
 */
function MultiRequest_applyToRange(
  ss,
  range_DadosJson,
  range_Status,
  range_Erro,
  range_Info,
  func_Apply,
  func_Trigger,
  func_Finally,
  num_MaxExecutions = 1,
  num_MaxTentativas = 1,
  time_MaxWaitLock = 3*60*1000,
  depurar = 0
){
  function DepurarLog(nivel, msg){
    if(depurar >= nivel) Logger.log(msg)
  }
  let time_LimitAppsScript = 6*60*1000
  let msg_Status = {
    'apply': 'Executando função...',
    'ok': 'Sucesso',
    'error': 'Falha'
  }

  // Execução da instância
  try{
    let date_StartInstance = new Date()
    let num_Execution = 0
    DepurarLog(1, 'Iniciando Execuções...')
    while(num_Execution < num_MaxExecutions){
      if(new Date() - date_StartInstance >= time_MaxWaitLock) break

      let info_Execution = {
        'date_StartExecution': null,
        'time_Execution': null,
        'date_ApplyFunction': null,
        'return_Function': null,
        'qtd_Tentativa': 0
      }
      let date_StartExecution = new Date()
      
      let lock_Document = LockService.getDocumentLock()
      // Execução da linha disponível
      try{
        lock_Document.waitLock(time_MaxWaitLock)
        
        DepurarLog(1, 'Obtendo intervalos...')
        let values_Data = range_DadosJson.getValues()
        let values_Status = range_Status.getValues()
        DepurarLog(2, `Valores dos dados:\n${values_Data}\n\nValores dos status:\n${values_Status}`)
        
        // Verifique a consistência dos intervalos
        if(values_Data.length != values_Status.length){
          throw new Error(
            'O tamanho do intervalo que contém os dados a serem consumidos e o status de cada execução deve possuir o mesmo tamanho'
          )
        }

        // Busque a linha disponível para realizar a requisição
        DepurarLog(1, `Buscando linha disponível para execução da função`)
        let num_RowCurrent = 1
        let i = 0
        while(i < values_Data.length){
          let status = values_Status[i][0]
          let dados = values_Data[i][0]

          // Condicional para linha disponível
          if(status == '' && dados != ""){

            num_Execution ++
            num_RowCurrent = i + 1
            let cell_CurrentStatus = range_Status.getCell(num_RowCurrent, 1)
            let cell_CurretInfo = range_Info.getCell(num_RowCurrent, 1)
            let value_Info = cell_CurretInfo.getValue()
            if(value_Info != ''){
              value_Info = JSON.parse(value_Info)
              info_Execution.qtd_Tentativa = value_Info['qtd_Tentativa']
            }
            info_Execution.qtd_Tentativa++
            info_Execution.date_StartExecution = date_StartExecution
            cell_CurretInfo.setValue(JSON.stringify(info_Execution, null, 2))
            
            // Atualiza o status da linha
            cell_CurrentStatus.setValue(msg_Status.apply)
            SpreadsheetApp.flush()
            lock_Document.releaseLock()
            
            // Aplica a função
            try{
              let dados_Json = JSON.parse(dados)
              DepurarLog(2, `Dados na linha disponível: ${dados_Json}`)
              
              DepurarLog(1, 'Aplicando função...')
              info_Execution.date_ApplyFunction = new Date()
              info_Execution.return_Function = func_Apply(dados_Json)
              
              cell_CurrentStatus.setValue(msg_Status.ok)
            }catch(error_ApplyFunc){
              let error_stack = error_ApplyFunc.stack
              DepurarLog(1, `Erro na execução da função:\n${error_stack}`)
              if(info_Execution.qtd_Tentativa < num_MaxTentativas){
                cell_CurrentStatus.setValue('')
              }else{
                cell_CurrentStatus.setValue(msg_Status.error)
              }
              range_Erro.getCell(num_RowCurrent, 1).setValue(error_stack)
              throw new Error(error_ApplyFunc.stack)
            }finally{
              info_Execution.time_Execution = new Date() - info_Execution['date_StartExecution']
              cell_CurretInfo.setValue(JSON.stringify(info_Execution, null, 2))
              break
            }
          }

          i++
        }

        // Condicional para parada das execuções e exclusão dos acionadores
        if(i == values_Data.length){
          DepurarLog(1, 'A função foi aplicada em todas as linhas, verificando se há execuções ocorrendo...')

          let has_Exetucion = false
          let i = 0
          let data_Info = range_Info.getValues()
          while(i < values_Data.length){
            let status = values_Status[i][0]
            let dados = values_Data[i][0]
            let info = data_Info[i][0]

            if(status != msg_Status.ok && status != msg_Status.error && dados != ""){
              DepurarLog(1, `Há execuções ocorrendo na linha ${i+1}`)
              has_Exetucion = has_Exetucion || true
              
              let dic_Info = JSON.parse(info)
              let date_Now = new Date()
              let date_Info = new Date(dic_Info['date_StartExecution'])
              DepurarLog(2, `Informações da execução:\n${info}\n\nData agora: ${date_Now}\n\nData da informação: ${date_Info}`)
              
              if(date_Now - date_Info >= time_LimitAppsScript){
                DepurarLog(1, `Esta execução ultrapassou o tempo limite, reiniciando o status dessa linha...`)
                range_Status.getCell(i+1, 1).setValue('')
              }
              
              break
            }
            i++
          }
          
          // Eliminação dos acionadores
          if(!has_Exetucion){
            let name_Func = func_Trigger.name
            DepurarLog(1, `Deletendo acionadores da função ${name_Func} ...`)
            ss.toast(`Todas as requisições foram realizadas, deletando acionadores...`)
            Triggers_Delete(name_Func)
            if(func_Finally != null) func_Finally()
          }

          // Parada das execuções
          break
        }
      }catch(error_CurrentExecutionRow){
        throw new Error(error_CurrentExecutionRow.stack)
      }finally{
        lock_Document.releaseLock()
      }
    }
  }catch(error_CurrentExecution){
    throw new Error(error_CurrentExecution.stack)
  }
}

/**
 * A função limpa as planilhas auxiliares e a planilha de controle da requisição (colunas de status). Depois a função cria os acionadores de determinada função, a ser passada como argumento.
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss - A spreadsheet que manuseia a API;
 * @param {Function} func - A função cujo acionadores serão criados;
 * @param {SpreadsheetApp.Range} range_DadosAc - Uma célula onde será colocada os dados do acionamento;
 * @param {SpreadsheetApp.Range} range_Status - Um intervalo coluna, em que cada célula contém o status de execução da linha. Este deve possuir o mesmo tamanho que o intervalo que contém os dados;
 * @param {SpreadsheetApp.Range} range_Erro - Um intervalo coluna, em que cada célula uma mensagem de erro com seus detalhes, caso o erro ocorra;
 * @param {SpreadsheetApp.Range} range_Info - Um intervalo coluna, em que cada célula o código colocará informações sobre a requisição;
 * @param {Number} n - O número de acionadores a serem criados
 * @param {Number} depurar - Nível de depuração para a função. Por padrão este argumento é 0. Quando diferente de zero, a função imprimi os log's para depuração, tendo os seguintes níveis de depuração:
 *    - 1: Mensagens principais
 */
function MultiRequest_Acionar(
  ss,
  func,
  range_DadosAc,
  range_Status,
  range_Erro,
  range_Info,
  n,
  depurar = 0
){
  function DepurarLog(nivel, msg){
    if(depurar >= nivel) Logger.log(msg)
  }
  // Verifique se não está ocorrendo requisições da API
  DepurarLog(1, 'Verificando se não está ocorrendo nenhuma requisição no momento...')
  if(Triggers_Match(func.name) == 1){
    let msg = "Espere até que as requisições sejam concluídas para criar novas requisições."
    try{
      SpreadsheetApp.getUi().alert(msg);    
    }finally{
      throw new Error(msg)
    }
  }

  range_DadosAc.setValue(new Date())

  // Limpe a planilha
  DepurarLog(1, 'Limpando a planilha...')
  ss.toast(`Limpando planilha de requisições...`)
  let documentProperties = PropertiesService.getDocumentProperties()
  documentProperties.deleteAllProperties()
  range_Status.clearContent()
  range_Erro.clearContent()
  range_Info.clearContent()
  SpreadsheetApp.flush()

  // Crie os acionadores
  DepurarLog(1, 'Criando acionadores...')
  if(n == 0){
    func()
  }else{
    ss.toast(`Criando acionadores...`)
    for(let i = 0; i < n; i++){
      let t1 = new Date()
      Trigger_Create(func.name)
      let t2 = new Date() - t1
      Utilities.sleep((60000/n)-t2)
    }
    let msg = `Os acionadores foram criados com sucesso. Após fechar este pop-up, acompanhe o status de cada uma das requisições.`
    try{
      SpreadsheetApp.getUi().alert(msg)
    }catch(error){
      ss.toast(msg)
    }
  }
}

/**
 * Obtém os intervalos nomeados de uma spreadsheet e os retorna como um dicionário, em que as chaves do dicionário são os nomes dos intervalos nomeados e as propriedades de cada chave são um objeto de intervalo.
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss - A spreadsheet a obter os intervalos nomeados
 * 
 * @returns {Object}
 */
function NamedRanges_Get(ss){
  let array_NamedRanges = ss.getNamedRanges();
  let dic_Ranges = {};
  for(let obj_Range of array_NamedRanges){
    let name = obj_Range.getName();
    let range = obj_Range.getRange();
    dic_Ranges[name] = range;
  }
  return dic_Ranges;
}


/**
 * Transforma uma SpreadSheet de tabelas aninhadas em um dicionário aninhado;
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss - A Spreadsheet que contém o esquema de aninhamento de tabelas
 * @param {String} key_Start - A chave de início, sobre o qual o aninhamento será construído
 * 
 */
function SheetNested_Json(ss, key_Start){
  sheet_Doc = ss.getSheetByName('DOC');
  // Transforme a planilha de documentação em um dicionário aninhado com profundidade de dois níveis;
  let data_Doc = sheet_Doc.getDataRange().getValues();
  data_Doc.shift();
  let dic_TipoData = {};
  for(let lin of data_Doc){
    let name_Sheet = lin[0];
    let arg = lin[1];
    let tipo = lin[2];

    if(!dic_TipoData[name_Sheet]) dic_TipoData[name_Sheet] = {};
    dic_TipoData[name_Sheet][arg] = tipo;
  }

  /**
   * Constroi um dicionário aninhado, a partir de um critério de recursão.
   * - A recursão acontece se o valor retornado pela função TipoData pertencer a uma lista especifica, definida no escopo da função.
   * 
   * @param {SpreadsheetApp.Sheet} name_Sheet - O nome da planilha a ser processada;
   * @param {String} saida - A referência superior a ser passada para o próximo nível de aninhamento;
   */
  function SheetAnin_DicAnin(name_Sheet = 'Pai', saida = key_Start){
    //Logger.log('Processando o planilha: ' + name_Sheet);
    let sheet = ss.getSheetByName(name_Sheet);
    
    // Construa um array de dicionários, passando linha a linha da planilha
    let array_DicsNow = [];    
    let data = sheet.getDataRange().getValues();
    let headers = data.slice(0,1)[0];
    data.shift();
    for(let i = 0; i< data.length; i++){
      let row = data[i];
      let chegada = row[0]
      
      if(chegada != saida) continue;

      // Ignore as primeiras duas colunas para construir o dicionário
      let dic_Row = {};
      for(let j = 2; j < row.length; j++){
        let key = headers[j];
        let tipo = dic_TipoData[sheet.getName()][key];
        let value = row[j];


        // Adicione um critério para recursão antes de adicionar o valor ao dicionário referente a linha atual
        let name_SheetProc  = name_Sheet + '_' + key;
        let saida_Proc = row[1]
        if(tipo == 'Array.Object'){
          // Obtenha o dicionário referente a próxima página
          if(value == 0 || value == ''){
            dic_Row[key] = [{}];
          }else{
            dic_Row[key] = SheetAnin_DicAnin(name_SheetProc, saida_Proc)
          };
        }else if(tipo == 'Object'){
          if(value == 0 || value == ''){
            dic_Row[key] = {};
          }else{
            dic_Row[key] = SheetAnin_DicAnin(name_SheetProc, saida_Proc)[0];
          }
        }else if(tipo == 'Array'){
          dic_Row[key] = [];
          let array_Dics = SheetAnin_DicAnin(name_SheetProc, saida_Proc);
          //Logger.log(JSON.stringify(array_Dics,null,2));
          for(let dic of array_Dics){
            dic_Row[key].push(Object.values(dic)[0]);
          }
        }else{
          dic_Row[key] = value;
        }
      }
      //Logger.log('Dicionário da linha:\n' + JSON.stringify(dic_Row,null,2));
      array_DicsNow.push(dic_Row)
    }
    //Logger.log('Dicionário da página:\n' + JSON.stringify(array_DicsNow, null, 2))
    return array_DicsNow;
  }

  let dic_Final = SheetAnin_DicAnin()[0];
  //Logger.log(JSON.stringify(dic_Final,null,2))
  return dic_Final;
}

/**
 * A função exporta um array de dados em um intervalo. É esperado que não haja repetição de elementos, e que o range seja um range coluna.
 * 
 * @param {SpreadsheetApp.Range} range - O range coluna a ser colocado os dados
 * @param {Array} array - O array de dados, sem repetições
 * @param {Number} depurar - Nível de depuração para a função. Por padrão este argumento é 0. Quando diferente de zero, a função imprimi os log's para depuração, tendo os seguintes níveis de depuração:
 *    - 1: Mensagens principais
 */
function ArrayToRange(range, array, depurar = 0){
  function DepurarLog(nivel, msg){
    if(depurar >= nivel) Logger.log(msg)
  }

  let dic_Keys = {}
  for(let key of array){
    dic_Keys[key] = key
  }
  DepurarLog(1, 'Obtendo semáforo')
  let lock_Script = LockService.getDocumentLock()
  try{
    DepurarLog(1, 'Esperando semáforo')
    lock_Script.waitLock(30000)

    DepurarLog(1, 'Removendo valores que já existem no intervalo')
    let data_SheetTransf = range.getValues()
    for(let i = data_SheetTransf.length - 1; i >= 0; i--){
      let key = data_SheetTransf[i][0]
      if(key != '' && dic_Keys.hasOwnProperty(key)){
        delete dic_Keys[key]
      }
    }

    let keys = Object.keys(dic_Keys)
    let num_Keys = keys.length
    let i = 0
    DepurarLog(1, 'Adicionando valores inéditos no intervalo')
    for(let j = 0; j < data_SheetTransf.length; j++){
      if(i == num_Keys) break
      if(data_SheetTransf[j][0] == ''){
        range.getCell(j+1,1).setValue(keys[i])
        i++
      }
    }
    SpreadsheetApp.flush()
  }catch(error){
    throw new Error(error.stack)
  }finally{
    DepurarLog(1, 'Liberando semáforo')
    lock_Script.releaseLock()
    SpreadsheetApp.flush()
  }
}

/**
 * A função recebe um array bidimensional, no formato de um dataframe, e exporta para uma planilha, passada como objeto num argumento da função, a partir de uma determinada linha. A função espera que a primeira linha da planilha seja o cabeçário dos dados. A função exporta coluna a coluna, não substituindo os valores do cabeçário, caso haja. Se houver uma coluna na planilha que não corresponde a uma coluna do dataFrame, esta coluna será limpa. Caso haja uma coluna no dataFrame que não exista na planilha, uma coluna será adicionada à planilha. Nesta adição é utilizado o serviço de semafóros do AppsScripts para evitar superposição de multuiplas instâncias.
 * 
 * @param {Array} dataFrame - Um array bidimensional de dados, em que a primeira linha corresponde ao cabeçário dos dados
 * @param {SpreadsheetApp.Sheet} sheet_Exp - A planilha onde os dados serão colocados.
 * @param {Number} row_Start - A linha a partir da qual começar a exportação
 */
function ArrayToSheet(dataFrame, sheet_Exp, row_Start){
  let data_HeaderDataFrame = dataFrame[0]
  let range_HeaderSheet = sheet_Exp.getRange('A1:1')

  // Atualiza o cabeçário, caso necessário
  let lock_Document = LockService.getDocumentLock()
  try{
    lock_Document.waitLock(30000)
    let data_HeaderSheet = range_HeaderSheet.getValues()[0]
    for(let name of data_HeaderDataFrame){
      let index = data_HeaderSheet.indexOf(name)
      if(index == -1){
        sheet_Exp.getRange(1, sheet_Exp.getLastColumn() + 1, 1, 1).setValue(name)
      }
    }
  }catch(e){
    throw new Error(e.message)
  }finally{
    SpreadsheetApp.flush()
    lock_Document.releaseLock()
  }


  // Obtenção da posição nas colunas do dataFrame
  let data_HeaderSheet = sheet_Exp.getRange('A1:1').getValues()[0]
  let positions = []
  for(let name of data_HeaderSheet){
    let index = data_HeaderDataFrame.indexOf(name)
    positions.push(index)
  }

  // Construa um array bidimensional na ordem correta
  let data_Ordered = []
  for(let i = 1; i < dataFrame.length; i++){
    let lin_Ordered = []
    for(let num_Col of positions){
      if(num_Col == -1){
        lin_Ordered.push('')
      }else{
        lin_Ordered.push(dataFrame[i][num_Col])
      }
    }
    data_Ordered.push(lin_Ordered)
  }

  // Exporte os dados
  let num_Rows = dataFrame.length - 1
  sheet_Exp.getRange(row_Start, 1, num_Rows, data_Ordered[0].length).setValues(data_Ordered)


  /* Exportação coluna a coluna

  for(let i = 0; i < positions.length; i++){
    let num_Col = positions[i]

    if(num_Col == -1){
      // Limpa a coluna, caso o dataFrame não contenha essa coluna
      sheet_Exp.getRange(row_Start, i + 1, num_Rows, 1).clearContent()
    }else{
      // Obtenha um array bidimensional da coluna
      let data_Col = []
      for(let j = 1; j < dataFrame.length; j++){
        data_Col.push([dataFrame[j][num_Col]])
      }

      // Exporte a coluna
      sheet_Exp.getRange(row_Start, i + 1, num_Rows, 1).setValues(data_Col)
    }
  }
  */
}

/**
 * Cria uma duplicada de uma sheet numa Spreadsheet, copiando as suas proteções. A função retorna o objeto da planilha criada.
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss - A spreadsheet que receberá a duplicata
 * @param {SpreadsheetApp.Sheet} sheet - A sheet que será duplicada
 */
function Sheet_Duplicate(ss, sheet){
  let sheet_New = sheet.copyTo(ss);

  // Copie as proteções da folha 'MODELO' para a nova folha
  let protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (let p of protections) {
    let rangeNotation = p.getRange().getA1Notation();
    let newProtection = sheet_New.getRange(rangeNotation).protect();
    newProtection.setDescription(p.getDescription());
    newProtection.setWarningOnly(p.isWarningOnly());
    if (!p.isWarningOnly()) {
      newProtection.removeEditors(newProtection.getEditors());
      newProtection.addEditors(p.getEditors());
    }
  }

  // Copie as proteções de página da folha 'MODELO' para a nova folha
  let sheetProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (sheetProtections.length > 0) {
    let p = sheetProtections[0];
    let newProtection = sheet_New.protect();
    newProtection.setDescription(p.getDescription());
    newProtection.setWarningOnly(p.isWarningOnly());
    if (!p.isWarningOnly()) {
      newProtection.removeEditors(newProtection.getEditors());
      newProtection.addEditors(p.getEditors());
    }

    // Copie as exceções de proteção da folha 'MODELO' para a nova folha
    let unprotectedRanges = p.getUnprotectedRanges();
    let newUnprotectedRanges = [];
    for (let r of unprotectedRanges) {
      let rangeNotation = r.getA1Notation();
      newUnprotectedRanges.push(sheet_New.getRange(rangeNotation));
    }
    newProtection.setUnprotectedRanges(newUnprotectedRanges);
  }
  return sheet_New
}
