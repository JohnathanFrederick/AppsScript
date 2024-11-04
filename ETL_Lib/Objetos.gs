/**
 * Obtém um dicionário aninhado, e retorna um dicionário desaninhado. Neste, cada chave é um caminho possível no dicionário aninhado, e cada propriedade é um array que contém os dicionários puros que estavam presentes no dicionário aninhado.
 * 
 * @param {Object} dic_Nested - Entre com um aninhamento de arrays de dicionários.
 * @param {Object} dic_PathPropagate - Entre com um objeto, em que suas chaves são o caminho do aninhamento, e o valor associado à chave é um array de propriedades a serem propagadas para o nível inferior.
 */
function Obj_Unnest(dic_Nested, dic_PathPropagate){
  let dic_Unnested = {};
  /**
   * Aplica a transforma no aninhamento
   * 
   * @param {Object} dic_Nested - Aninhamento de arrays de dicionários a ser desaninhado;
   * @param {string} path_Current - O caminho para o nível atual;
   * @param {number} depth - Argumento para recursão, identifica a profundidade da recursão;
   * @param {Object} dic_Upper - Um dicionário do nível superior, contendo valores a serem propagados.
   */
  function ObjAn_Obj(dic_Nested, path_Current = 'Pai', depth = 0, dic_Upper = {}){
    if(!dic_Nested) return;

    /**
     * Estrutura de recursão
     * 
     * @param {Object} dic_Nested - O dicionário a qual aplicar a transformação.
     */
    function HandleObject(dic_Nested){
      let dic_Pure = {...dic_Upper};
      let dic_Prop = {...dic_Upper};
      
      // Obtenha os dados puros e os que serão propagados
      for(let key in dic_Nested){
        let data = dic_Nested[key];
        if(typeof data !== 'object'){
          dic_Pure[key] = data;
          
          // Verifique se esse aninhamento possui valores a serem propagados
          if(path_Current in dic_PathPropagate){
            if(dic_PathPropagate[path_Current].includes(key)){
              dic_Prop[path_Current + "." + key] = data;
            }
          }
        }
      }

      // Realize a recursão
      for(let key in dic_Nested){
        let data = dic_Nested[key];
        if(typeof data === 'object'){
          let path_Next = path_Current + "." + key;
          ObjAn_Obj(data, path_Next, depth + 1, dic_Prop);
        }
      }

      // Adicione o dicionário puro ao dicionário desaninhado
      if(!(path_Current in dic_Unnested)) dic_Unnested[path_Current] = []
      dic_Unnested[path_Current].push(dic_Pure)
    }

    if(Array.isArray(dic_Nested)){
      for(let dic of dic_Nested){
        HandleObject(dic);
      }
    }else if(typeof dic_Nested == 'object'){
      HandleObject(dic_Nested);
    }
  }

  ObjAn_Obj(dic_Nested);
  return dic_Unnested
}

/**
 * A função transforma um array de dicionários em um array bidimensional, no formato de um dataframe
 * 
 * @param {Array} array_Dic - Um array de dicionários, a ser transformado em um array bidimensional
 * 
 */
function ArrayDic_ToDataFrame(array_Dic){
  let header = []
  let data = []
  for(let dic of array_Dic){
    let lin = []
    for(let key_Header of header){
      if(key_Header in dic){
        lin.push(dic[key_Header])
        delete dic[key_Header]
      }else{
        lin.push(null)
      }
    }
    for(let key in dic){
      header.push(key)
      lin.push(dic[key])
    }
    data.push(lin)
  }
  return [header].concat(data)
}
