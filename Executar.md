~~~javascript
/**
 * A função limpa as planilhas auxiliares e a planilha de controle da requisição (colunas de status). Depois a função cria os acionadores de determinada função, a ser passada como argumento.
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss - A spreadsheet que manuseia a API;
 * @param {String} name_Func - O nome da função cujo acionadores serão criados
 */
function CriarAcionadores(ss, name_Func = 'Get'){
  let dic_Ranges = Util_GetNamedRanges(ss);
  let range_StatusProc = dic_Ranges['Status_Proc'];
  let range_NumAc = dic_Ranges['Num_Acionadores'];
  let range_Method = dic_Ranges['Method'];
  let range_UltAc = dic_Ranges['Ultimo_Acionamento'];
  let range_Proc = dic_Ranges['Processamento'];
  
  // Verifique se não está ocorrendo requisições da API
  Logger.log('Verificando se não está ocorrendo nenhuma requisição no momento...')
  let status = range_StatusProc.getValue();
  if(status != "Requisições completas"){
    if(Util_VerifiqueTriggers(name_Func) == 1){
      ss.toast("Espere até que as requisições sejam concluídas para criar novas requisições.");
      return;
    }
  }

  // Limpe a página PAINEL
  Logger.log('Limpando a página PAINEL...')
  let documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.deleteAllProperties();
  let method = range_Method.getValue();
  let lenght_Cleaning = 5;
  if(method == 'POST') lenght_Cleaning = 4;
  let col_StartClean = range_Proc.getNumColumns() - lenght_Cleaning + 1;
  range_Proc.offset(0,col_StartClean - 1, range_Proc.getNumRows(),lenght_Cleaning).clearContent();
  SpreadsheetApp.flush();

  // Crie os acionadores e informe na planilha...
  Logger.log('Criando acionadores...')
  range_UltAc.setValue(new Date());
  range_StatusProc.setValue("Criando acionadores...");
  SpreadsheetApp.flush();
  let n = range_NumAc.getValue();
  for(let i=0; i<n; i++){
    let t1 = new Date();
    Util_CreateTriggers(name_Func);
    let t2 = new Date() - t1;
    Utilities.sleep((60000/n)-t2);
  }

  ss.toast('Acionadores criados com sucesso, suas requisições estão em processamento.')
}
/**
 * Realiza uma única requisição da API UAU, de qualquer endpoint, processando a planilha para obter os parâmetros necessários.
 * A função possui tratamento de semafóros para lidar com múltiplas instâncias, permitindo multiplas requisições da API. 
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss - A spreadsheet a ser processada, destinada ao manuseio da API;
 * @param {String} name_FuncAc - O nome da função que está processando as requisições. A função deletará os acionadores com este nome, quando não houver mais linhas a serem processadas;
 * @param {String} [url_Api] - A url para a requisição da API.
 * 
 */
function Executar_Request(ss, name_FuncAc, url_Api, params_base = {}, headers_base, map_Payload = false) {
  let date_Start = new Date();
  let dic_Ranges = Util_GetNamedRanges(ss);

  // Parâmetros de Requisição
  let range_ParamsRequest = dic_Ranges['Request_Params'];
  let range_Method = dic_Ranges['Method'];
  let data_Params = range_ParamsRequest.getValues();
  data_Params = data_Params.filter(function(row){return !row.every(function(cell){return cell === ""})})
  let num_Params = data_Params.length;
  let method = range_Method.getValue();
  
  // Processamento das Requisições
  let range_StatusProc = dic_Ranges['Status_Proc'];
  let range_UltAc = dic_Ranges['Ultimo_Acionamento'];
  let range_Proc = dic_Ranges['Processamento'];
  let time_UltSoli = range_UltAc.getValue();

  // Transformação e Exportação
  if(method == 'GET'){
    var range_TranfExp = dic_Ranges['Transformacao_Exportacao'];
  }

  // Função para a finalização do processamento
  function Finally(){
    // Informe o processo de finalização
    range_StatusProc.setValue("Finalizando o processo...");
    SpreadsheetApp.flush();


    // Pare de realizar requisições
    Util_DeleteTriggers(name_FuncAc);

    // Libere todos os bloqueios
    let lock_Document = LockService.getDocumentLock();
    lock_Document.releaseLock();

    // Informe a finalização das requisições
    range_StatusProc.setValue("Requisições completas");
    SpreadsheetApp.flush();

    if(method == 'GET'){
      let documentProperties = PropertiesService.getDocumentProperties();
      // Delete as linhas sobressalentes
      let data_TransExp = range_TranfExp.getValues();
      for(let i = 0; i < data_TransExp.length; i++){
        let key_Data = data_TransExp[i][0];
        let url = data_TransExp[i][2];
        let name_Sheet = data_TransExp[i][3];
        if(!!key_Data && !!url){
          let key = "Row." + key_Data;
          let lin_start = Math.floor(parseFloat(documentProperties.getProperty(key)))
          let sheet_Destino = SpreadsheetApp.openByUrl(url).getSheetByName(name_Sheet);
          let num_DeleteRowns = sheet_Destino.getMaxRows()-lin_start+1;
          if(num_DeleteRowns > 0){
            sheet_Destino.deleteRows(lin_start,num_DeleteRowns);
          }
          SpreadsheetApp.flush();
        }
      }
      // Limpe as propriedades do documento
      documentProperties.deleteAllProperties();
    }
  }

  // Obtenha os parâmetros para o processamento das requisições
  /**
   * Obtém os dados da planilha e os formata no formato exigido para a requisição da API.
   * 
   * @param {SpreadsheetApp.Spreadsheet} ss - A spreadsheet a ser processada para obter os parâmetros de requisição.
   * - É necessário que ela esteja no formato padrão que criamos para o processamento da API.
   * @param {SpreadsheetApp.Sheet} sheet_Painel - Página (sheet) a ser processada;
   */
  let allParams = {};
  function Params_(){
    let requestParams = {};
    let i = 0;
    let data_Proc = range_Proc.getValues();
    //Logger.log(data_Proc);
    data_Proc = data_Proc.filter(function(row){return row[0] != ''})
    if(data_Proc.length < 1) data_Proc.push(['']);
    //Logger.log(JSON.stringify(data_Proc, null, 2))

    // Percorra todas as linhas dos parâmetros de requisição, em busca de uma única linha que ainda não foi requisitada
    let t = 0;
    do{
      let row_Now = data_Proc[i];
      // Verifique o status da linha
      let status_LinRequest = row_Now[num_Params];

      // Conte as linhas que estão pendentes de processamento
      if(status_LinRequest != 'OK' && status_LinRequest != 'FALHA') t++;

      // Verifique se a linha atual irá ser processada
      let teste = status_LinRequest == '';
      if(num_Params > 0){
        teste = teste && row_Now[0] !== '';
      }
      if(teste){
        // Mude o status da linha, a fim de liberar o semáforo
        range_Proc.getCell(i + 1, num_Params + 1).setValue('OBTENDO...')
        SpreadsheetApp.flush();
  
        // Obtenha os dados
        for(let j = 0; j < num_Params; j++){
          let data = row_Now[j];
          let formatData = data_Params[j][1];
          let nameData = data_Params[j][0];
  
          if(formatData == 'ANINHAMENTO'){
            data = SheetToJson(ss, data);
            requestParams = Object.assign(requestParams, data);
          }else if(formatData == 'ENDPOINT'){
            url_Api = url_Api + data;
          }else if(formatData == 'JSON'){
            requestParams[nameData] = JSON.parse(data);
          }else{
            requestParams[nameData] = data;
          }

          allParams[nameData] = data;
        }
        break
      }
      i++;
    }while(i < data_Proc.length);
    
    let period_Proc = (new Date().getTime() - time_UltSoli.getTime());
    let calc = period_Proc/(60*1000)/(7 * data_Proc.length);
    Logger.log('Valor do teste temporal: ' + calc + '\nCaso esse valor se torne maior ou igual a 1, o processo será finalizado.')

    if(calc >= 1){
      Logger.log('Os acionadores serão deletados por excederem o tempo limite.');
      Finally();
      return
    }

    // Verifique se todas as linhas já foram processadas
    if(t == 0){
      Finally();
      return
    }else if(i == data_Proc.length){
      return;
    }

    range_StatusProc.setValue("Obtendo dados da iteração: " + i);

    let dic_Request = requestParams;
    if(!!params_base) dic_Request = Object.assign(params_base, dic_Request);

    return [dic_Request,i]
  }
  // Coloque um semáforo na planilha
  let lock_Document = LockService.getDocumentLock();
  lock_Document.waitLock(30000);
  try{
    let response_Params = Params_(ss);
    if(Array.isArray(response_Params)){
      var [dic_RequestParams,position] = response_Params;
      // Logger.log(JSON.stringify(dic_RequestParams,null,2));
    }else{
      Logger.log('Não conseguimos obter os parâmetros da página. Provavelmente todos os parâmetros já foram obtidos.')
      return
    }
  }catch(error){
    Logger.log(error.stack);
    throw new Error(error)
  }finally{
    lock_Document.releaseLock();
  }

  // Função para o tratamento de erros
  let range_StatusLinNow = range_Proc.getCell(position + 1, num_Params + 1)
  function DeuRuim(error){
    // Registre o erro na planilha
    let num_Tent = range_Proc.getCell(position + 1, num_Params + 3).getValue();
    if(num_Tent < 3){
      range_StatusLinNow.setValue("");
      range_Proc.getCell(position + 1, num_Params + 3).setValue(num_Tent + 1);
    }else{
      range_StatusLinNow.setValue("FALHA");
      range_Proc.getCell(position + 1, num_Params + 3).setValue(num_Tent + 1);
    }

    // Grave o erro no registro de execução
    let date_End = new Date();
    let period = date_End - date_Start;
    range_Proc.getCell(position + 1, num_Params + 2).setValue(error.message);
    range_Proc.getCell(position + 1, num_Params + 4).setValue(period);
    Logger.log('url de requisição: ' + url_Api);
    Logger.log('Pilha de chamada de erro:\n\n' + error.stack);
    Logger.log('Objeto de erro:\n\n' + JSON.stringify(error,null,2));
    throw new Error(error.message);
  }

  // Faça a requisição
  Logger.log('Requisitando...')
  /**
   * Realiza a requisição HTTP de uma URL da API do ClickUp.
   * 
   * @param {string} url - Entre com a url da API requisitada. A url deve estar entre aspas.
   * @param {Object} requestParams - Entre com os parâmetros da requisição. Estes devem ser colocados como Objects. Para mais informações consulte a documentação da API.
   * @param {String} method - O método de requisição HTTP, se omitido a função define este parâmetro como 'GET'.
   * 
   */
  function Request_(url,requestParams, method = 'GET'){
    let headers = headers_base;

    try{
      if(method == 'GET'){
        var data = JSON.parse(Request_ApiRestful(url,method,headers, requestParams))
      }else{
        Request_ApiRestful(url,method,headers, requestParams);
        return;
      }
    }catch(error){
      Logger.log(error.stack);
      throw new Error(error);
    }
    return data;
  }
  try{
    var response = Request_(url_Api,dic_RequestParams,method)
  }catch(error){
    DeuRuim(error);
  }

  // Transforme e Exporte os dados obtidos, para o método GET
  if(method == 'GET'){
    try{
      // Obtenha os parâmetros de transformação como um dicionário
      let data_TransExp = range_TranfExp.getValues();
      let dic_PropKeys = {};
      for(let lin of data_TransExp){
        let caminho = lin[0];
        let keys_Prop = lin[1];
        if(keys_Prop !== '' && caminho !== ''){
          try{
            dic_PropKeys[caminho] = JSON.parse(keys_Prop);
          }catch(e){
            Logger.log(e + '\n\nString que tentamos converter:\n\n' + keys_Prop);
          }
        }
      }

      // Aplique a transformação
      Logger.log('Transformando...')
      range_StatusLinNow.setValue("TRANSFORMANDO");
      SpreadsheetApp.flush();
      let dic_Data = Transf_DicAninDic(response, dic_PropKeys);

      // Exporte os dados obtidos
      Logger.log('Exportando...')
      range_StatusLinNow.setValue("EXPORTANDO");
      SpreadsheetApp.flush();

      for(let key in dic_Data){
        let data_TransExp = range_TranfExp.getValues();
        let data = dic_Data[key]
        
        function ExportarDados(lin){
          // Critério para exportação
          let url = lin[2];
          if(url == '') return

          // Adicione o valor a ser mapeado aos dados
          let obj_map = {};
          obj_map['Data de Requisição'] = time_UltSoli;
          if(map_Payload){
            obj_map = Object.assign(obj_map, allParams)
            Logger.log(JSON.stringify(obj_map, null, 2));
          }
          let header_Data = data[0].concat(Object.keys(obj_map));
          let newData = data.slice(1).map(row => row.concat(Object.values(obj_map)))
          newData.unshift(header_Data);
          data = newData;

          // Adiciona um semáforo na planilha, para que não haja sobreposição de dados na planilha de destino
          let lock = LockService.getScriptLock();
          lock.waitLock(200000);

          // Adiciona colunas inéditas ao cabeçário da planilha de destino
          let name_Sheet = lin[3];
          let sheet_Dest = SpreadsheetApp.openByUrl(url).getSheetByName(name_Sheet);
          let num_ColSheetDest = sheet_Dest.getLastColumn();
          let hearder_DestSheet = [];
          if(num_ColSheetDest > 0) hearder_DestSheet = sheet_Dest.getRange(1,1,1,sheet_Dest.getLastColumn()).getValues()[0];
          hearder_DestSheet = hearder_DestSheet.filter(function(cell){return cell !== ''})
          header_Data.forEach((h) => {
            if(!hearder_DestSheet.includes(h)){
              hearder_DestSheet.push(h);
            }
          });
          sheet_Dest.getRange(1,1,1,hearder_DestSheet.length).setValues([hearder_DestSheet]);
          SpreadsheetApp.flush();

          // Reorganize os dados para corresponder ao cabeçalho da planilha de destino
          let data_Ordered = data.slice(1).map(row => {
            let obj = {};
            header_Data.forEach((h, i) => obj[h] = row[i]);
            return hearder_DestSheet.map(h => obj[h] || '')
          })

          // Obtenha a linha inicial de exportação
          let name_KeyPropertie = "Row." + key;
          let documentProperties = PropertiesService.getDocumentProperties();
          let lin_start = parseInt(documentProperties.getProperty(name_KeyPropertie),10) || 2;

          // Remova qualquer filtro que esteja sendo aplicado à planilha
          if(sheet_Dest.getFilter()) sheet_Dest.getFilter().remove();

          // Exporte os dados
          try{
            sheet_Dest.getRange(lin_start,1,data_Ordered.length,data_Ordered[0].length).setValues(data_Ordered);
            // Atualiza a propriedade a propriedade da planilha
            documentProperties.setProperty(name_KeyPropertie,lin_start + data_Ordered.length)
          }catch(e){
            Logger.log('Erro ao tentar exportar os dados com o caminho: '+ key + '\nErro obtido: ' + e);
            Logger.log('Dados que tentamos exportar:\n\n' + JSON.stringify(data_Ordered,null,2))
          }finally{
            SpreadsheetApp.flush();
            lock.releaseLock();
          }
        }

        // Procure por esta chave em todos os caminhos
        let t = 0;
        for(let i = 0; i < data_TransExp.length; i++){
          let lin = data_TransExp[i]
          let caminho = lin[0];
          if(caminho == key){
            ExportarDados(lin);
            t++
            break;
          }
        }

        // Caso não exista, insira um novo caminho
        if(t == 0){
          for(let i = 0; i < data_TransExp.length; i++){
            let lin = data_TransExp[i];
            let caminho = lin[0];

            if(caminho == ''){
              range_TranfExp.getCell(i + 1, 1).setValue(key);
              SpreadsheetApp.flush();
              ExportarDados(lin);
              break;
            }
          }
        }
      }

      // Grave na planilha o sucesso
      let n = 0;
      for(let key in dic_Data){
        if(dic_Data[key].length>0){
          n = n - 1;
        }
        n = n + dic_Data[key].length;
      }
      range_Proc.getCell(position + 1, num_Params + 5).setValue(n);
    }catch(error){
      DeuRuim(error);
    }
  }
  let date_End = new Date();
  let period = date_End - date_Start;
  range_Proc.getCell(position + 1, num_Params + 4).setValue(period);
  range_StatusLinNow.setValue('OK');
  SpreadsheetApp.flush();
}
~~~
