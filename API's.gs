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
function Api_Request(url, method = 'GET', headers, params) {
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
