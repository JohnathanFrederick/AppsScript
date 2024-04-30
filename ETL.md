
~~~javascript
/**
 * Transforma um aninhamento de arrays de dicionários em um único dicionário de arrays bidimensionais, agrupando os arrays com mesmo nível e caminho de profundidade. 
 * 
 * @param {Object} content - Entre com um aninhamento de arrays de dicionários.
 * @param {Object} keyRefs - Entre com um objeto, em que suas chaves são o caminho do aninhamento, e o valor associado à chave é um array de propriedades a serem propagadas para o nível inferior.
 */
function Transf_DicAninDic(content, keyRefs = {'':[]}){
  /**
   * A função adiciona um dicionário puro (sem estrutura de aninhamento) a um dicionário de arrays bidimensionais, em que o array bidimensional que irá acumular os dados do dicionário puro é selecionado através de uma chave, passada como argumento
   */
  function Add_DicPureToDicArrayBi(dic_Final,dic_Pure,indKey){
    // Obtenha o array que acumula os dados
    if(!dic_Final.hasOwnProperty(indKey)) dic_Final[indKey] = [Object.keys(dic_Pure)];
    let array_Acum = dic_Final[indKey];
    let cab_ArrayAcum = array_Acum[0];
    
    // Organize a linha que será adicionada a este array
    let newLin = [];
    let keys_YetCons = [];
    for(let key_Cab of cab_ArrayAcum){
      if(key_Cab in dic_Pure){
        newLin.push(dic_Pure[key_Cab]);
        keys_YetCons.push(key_Cab);
      }else{
        newLin.push("");
      }
    }

    // Adicione a linha organizada
    dic_Final[indKey] = dic_Final[indKey].concat([newLin]);
    array_Acum = dic_Final[indKey]
    
    // Adicione as colunas inéditas, preenchendo com vazio as linhas que não possui essa coluna
    for(let key_Cons in dic_Pure){
      if(!keys_YetCons.includes(key_Cons)){
        if(!!dic_Pure[key_Cons]){
          dic_Final[indKey][array_Acum.length - 1].push(dic_Pure[key_Cons]);
        }else{
          dic_Final[indKey][array_Acum.length - 1].push('');
        }
        
        // Adicione a chave ao cabeçário
        dic_Final[indKey][0].push(key_Cons);

        // Preencha com vazio a coluna das linhas que não possuem a chave inédita
        for(let i = 1; i < array_Acum.length - 1; i++){
          dic_Final[indKey][i].push("");
        }
      }
    }
  }

  // Construa o dicionário
  let dic_Final = {};
  /**
   * Aplica a transforma no aninhamento
   * 
   * @param {Object} content - Aninhamento de arrays de dicionários a ser desaninhado;
   * @param {string} indKey - Argumento para recursão, propaga as chaves de um nível superior para o nível inferior;
   * @param {number} depth - Argumento para recursão, identifica a profundidade da recursão;
   * @param {Object} valsSup - Argumento para propagar as chaves e os valores de um nível superior para o nível inferior.
   */
  function ObjAn_Obj(content, indKey = 'Pai', depth = 0, valsSup = {}){
    /*
    Logger.log(
      'Entrando no Objeto: ' + indKey
      + '\n\nEste objeto é do tipo: ' + typeof(content)
      + '\n\nAs chaves que estão sendo propagadas são: ' + JSON.stringify(valsSup,null,2)
    );
    */
    
    // Critério de parada da recursão
    if(!content) return;

    /**
     * Estrutura de recursão
     * 
     * @param {Object} dic_DataAnin - O dicionário a qual aplicar a transformação.
     */
    function HandleObject(dic_DataAnin){
      let dic_Pure = {...valsSup};
      let dic_Prop = {...valsSup};
      
      // Obtenha os dados puros e os que serão propagados
      for(let key in dic_DataAnin){
        let data = dic_DataAnin[key];
        if(typeof data !== 'object'){
          dic_Pure[key] = data;
          
          // Verifique se esse aninhamento possui valore a serem propagados
          if(indKey in keyRefs){
            // Verifique se esta chave precisa ser propagada
            if(keyRefs[indKey].includes(key)){
              let key_Prod = indKey + "_" + key;
              dic_Prop[key_Prod] = data;
            }
          }
        }
      }

      // Realize a recursão
      for(let key in dic_DataAnin){
        let data = dic_DataAnin[key];
        if(typeof data === 'object'){
          dic_Pure[key] = "";
          let newKey = indKey + "." + key;
          ObjAn_Obj(data, newKey, depth+1, dic_Prop);
        }
      }

      // Adicione o dicionário puro obtido ao dicionário final
      Add_DicPureToDicArrayBi(dic_Final,dic_Pure,indKey);
    }

    if(Array.isArray(content)){
      for(let dic of content){
        HandleObject(dic);
      }
    }else if(typeof content == 'object'){
      HandleObject(content);
    }
  }

  ObjAn_Obj(content);
  return dic_Final
}

/**
 * Transforma uma SpreadSheet de tabelas aninhadas em um dicionário aninhado;
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss - A Spreadsheet que contém o esquema de aninhamento de tabelas
 * @param {String} key_Start - A chave de início, sobre o qual o aninhamento será construído
 * 
 */
function SheetToJson(ss, key_Start){
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
 * Realiza uma requisição HTTP para uma API RESTful e retorna os dados recebidos no formato String JSON.
 * A função suporta métodos HTTP personalizados, cabeçalhos e corpos de requisição.
 * 
 * @param {String} url - O URL da API onde a requisição será feita.
 * @param {String} [method='GET'] - O método HTTP da requisição (GET, POST, PUT, DELETE, etc.).
 *                                - Se não fornecido, 'GET' será usado por padrão.
 * @param {Object} [headers] - Os cabeçalhos da requisição HTTP. Deve seguir o formato:
 *                           - {'Content-Type': 'application/json', 'Authorization': 'Bearer seu_token'}
 * @param {Object} [params] - O corpo da requisição para métodos como POST ou PUT.
 *                           - Deve ser um objeto válido.
 * 
 * @returns {String} - Uma string JSON com os dados retornados pela API.
 * 
 * @throws {Error} - Lança um erro se a requisição falhar ou se a resposta não puder ser analisada.
 * 
 * @example
 * // Exemplo de requisição GET
 * var data = Request('https://api.exemplo.com/v1/recurso');
 * 
 * // Exemplo de requisição POST com cabeçalhos e corpo
 * var data = Request('https://api.exemplo.com/v1/recurso', 'POST', 
 *                    {'Content-Type': 'application/json', 'Authorization': 'Bearer seu_token'},
 *                    {'param1': 'valor1', 'param2': 'valor2'});
 */
function Request_ApiRestful(url, method = 'GET', headers, params) {
  let options = {
    method: method,
    headers: headers,
    muteHttpExceptions: true
  };
  if(params != null){
    if(Object.keys(params).length > 0){
      options.payload = JSON.stringify(params);
    }
  }
  try{
    var response = UrlFetchApp.fetch(url, options);
    if(response.getResponseCode() >= 400){
      throw new Error(response.getContentText())
    }
  }catch(error){
    Logger.log("Houve um erro ao processar a sua requisição. Verifique os dados mais detalhados do erro abaixo:\n\n"
      + "Os dados de requisição obtidos foram:\n\n" 
      + JSON.stringify(options,null,2) 
      + "\n\nO payload obtido foi:\n\n"
      + JSON.stringify(params, null, 2)
    );
    throw new Error(error);
  }
  return response
}
~~~
