/* var soma = require("./soma")
var sub = require("./sub")
var mult = require("./mult")
var div = require("./div")

console.log(`O valor da Soma é: ${soma}.`)
console.log(`O valor da Subtração é: ${sub}.`)
console.log(`O valor da Multiplicação é: ${mult}.`)
console.log(`O valor da Divisão é: ${div}.`) */

function insert(num){

    var numero = document.getElementById('resultado').innerHTML;
    document.getElementById('resultado').innerHTML = numero + num;
}
function clean()
{
    document.getElementById('resultado').innerHTML = "";
}
function back()
{
    var resultado = document.getElementById('resultado').innerHTML;
    document.getElementById('resultado').innerHTML = resultado.substring(0, resultado.length -1);
}
function calcular()
{
    var resultado = document.getElementById('resultado').innerHTML;
    if(resultado)
    {
        document.getElementById('resultado').innerHTML = eval(resultado);
    }
    else
    {
        document.getElementById('resultado').innerHTML = "Nada..."
    }
}
