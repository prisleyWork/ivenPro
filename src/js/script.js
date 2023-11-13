var workbook;
var itensCorretos = [];
var itensIncorretos = [];

fetch('BENS.xlsx')
.then(response => response.arrayBuffer())
.then(data => {
   var data = new Uint8Array(data);
   workbook = XLSX.read(data, {type: 'array'});
});

const anoAtualSpan = document.getElementById("anoAtual");
const dataAtual = new Date();
const ano = dataAtual.getFullYear();
anoAtualSpan.textContent = ano;

function formatInput(input) {
    var noLeadingZeros = input.replace(/^0+/, '');
    var formattedInput = noLeadingZeros.replace(/-\d$/, '');
    return formattedInput;
}

function translateCondition(condition) {
    var translations = {
        'BM': 'Bom',
        'AE': 'Anti-Econômico',
        'IR': 'Irrecuperável',
        'OC': 'Ocioso',
        'BX': 'Baixado',
        'RE': 'Recuperável'
    };
    
    return translations[condition] || condition;
}

function translateSituation(situation) {
    var translations = {
        'NI': 'Não encontrado no local da guarda',
        'NO': 'Normal'
    };
    
    return translations[situation] || situation;
}

document.getElementById('numero').addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        buscar();
    }
});

function buscar() {
    var inputField = document.getElementById('numero');
    var formattedInput = formatInput(inputField.value);
    var selectedRoom = document.getElementById('sala').value;
    
    if (selectedRoom === "") {
        alert("Por favor, selecione uma sala antes de buscar.");
        return;
    }
    
    var worksheet = workbook.Sheets[workbook.SheetNames[0]];
    var jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});
   
   for (var i = 0; i < jsonData.length; i++) {
       if (jsonData[i][0] == formattedInput || jsonData[i][2] == formattedInput) {
           var translatedCondition = translateCondition(jsonData[i][3]);
           var translatedSituation = translateSituation(jsonData[i][4]);
           
           var resultBackground = (jsonData[i][7].toLowerCase() == selectedRoom.toLowerCase()) ? '#0080005d' : '#f51515af';
           
           document.getElementById('resultado').style.backgroundColor = resultBackground;
           document.getElementById('resultado').innerHTML = 'Número de patrimônio: ' + jsonData[i][0] + '-' + jsonData[i][1] +'<br>' +
                                                            'Tipo: ' + jsonData[i][8] + '<br>' +
                                                            'Descrição: ' + jsonData[i][5] + '<br>' +
                                                            'Situação do Bem: ' + translatedSituation +'<br>' +
                                                            'Condição do Bem: ' + translatedCondition + '<br>' +
                                                            'Local: ' + jsonData[i][7] + '<br>' +
                                                            'Responsável: ' + jsonData[i][9];
           inputField.value = '';
           inputField.focus();

           if (jsonData[i][9].toLowerCase() == selectedRoom.toLowerCase()) {
               itensCorretos.push(jsonData[i]);
           } else {
               itensIncorretos.push(jsonData[i]);
           }

           return;
       }
   }
   
   document.getElementById('resultado').innerHTML = "Número não encontrado no banco de dados.";
   
   inputField.value = '';
   inputField.focus();
}
window.onload = function() {
    var select = document.getElementById("sala");
    var options = Array.prototype.slice.call(select.getElementsByTagName('option'));
    options.sort(function(a, b) {
        if (a.text > b.text) return 1;
        if (a.text < b.text) return -1;
        return 0;
    });
    while (select.firstChild) {
        select.removeChild(select.firstChild);
    }
    options.forEach(function(option) {
        select.appendChild(option);
    });

    document.getElementById('numero').focus();
};
function exibirLog() {
    console.log("Itens corretos:");
    console.log(itensCorretos);
    console.log("Itens incorretos:");
    console.log(itensIncorretos);
}
function atualizarItem(e) {
    var checkbox = e.target;
    var itemDiv = checkbox.parentElement;
    var itemNumber = itemDiv.firstChild.textContent.split(' ')[0];
    
    var itemArray = checkbox.checked ? itensCorretos : itensIncorretos;
    for (var i = 0; i < itemArray.length; i++) {
        if (itemArray[i][0] == itemNumber) {
            itemArray[i][10] = checkbox.checked ? 'Ocioso' : 'Não ocioso';
            break;
        }
    }
}

function Gerar() {
    var resultadoDiv = document.getElementById('resultado');
    resultadoDiv.innerHTML = '';

    var selectedRoom = document.getElementById('sala').value;

    var bensCorretosSet = new Set(itensCorretos.map(item => item[0] + '-' + item[1] + ' ' + item[8]));
    var bensIncorretosSet = new Set(itensIncorretos.map(item => item[0] + '-' + item[1] + ' ' + item[8] + ' origem: ' + item[7] + ' Responsável: ' + item[9]));

    var salaDiv = document.createElement('div');
    salaDiv.innerHTML = '<h2>Relação dos bens: ' + selectedRoom + '</h2>';
    resultadoDiv.appendChild(salaDiv);

    var totalBensVerificados = bensCorretosSet.size + bensIncorretosSet.size;
    var totalDiv = document.createElement('div');
     totalDiv.innerHTML = '<h2>Total de bens verificados: ' + totalBensVerificados + '</h2>';
    resultadoDiv.appendChild(totalDiv);

    var bensCorretosDiv = document.createElement('div');
    bensCorretosDiv.innerHTML = '<h2>Bens corretamente alocados (' + bensCorretosSet.size + '):</h2>';
    Array.from(bensCorretosSet).forEach(bem => {
        var bemDiv = document.createElement('div');
        bemDiv.innerHTML = bem + ' <input type="checkbox"> Ocioso <input type="checkbox"> Quebrado <input type="checkbox" onclick="outros(this)"> Não encontrado <input type="text" style="display: none"> <input type="checkbox"> Sem placa ';
        bensCorretosDiv.appendChild(bemDiv);
    });
    resultadoDiv.appendChild(bensCorretosDiv);

    var bensIncorretosDiv = document.createElement('div');
    bensIncorretosDiv.innerHTML = '<h2>Bens pertencentes a outros locais (' + bensIncorretosSet.size + '):</h2>';
    Array.from(bensIncorretosSet).forEach(bem => {
        var bemDiv = document.createElement('div');
        bemDiv.innerHTML = bem + ' <input type="checkbox"> Ocioso <input type="checkbox"> Quebrando <input type="checkbox"> Sem placa <input type="checkbox" onclick="outros(this)"> Outros <input type="text" style="display: none">';
        bensIncorretosDiv.appendChild(bemDiv);
    });
    resultadoDiv.appendChild(bensIncorretosDiv);
    function getCheckboxValues(div) {
        var checkboxes = div.querySelectorAll('input[type="checkbox"]');
        var values = [];
        checkboxes.forEach(function(checkbox) {
            if (checkbox.checked) {
                var label = checkbox.nextSibling.textContent.trim();
                var input = checkbox.nextSibling.nextSibling;
                if (input.style.display !== 'none') {
                    label += ' ' + input.value;
                }
                values.push(label);
            }
        });
        return values.join(' ');
    }
    
}

function removerCheckboxNaoMarcadas() {
    var resultadoDiv = document.getElementById('resultado');
    var divsDeItem = resultadoDiv.getElementsByTagName('div');

    for (var i = 0; i < divsDeItem.length; i++) {
        var divDeItem = divsDeItem[i];
        var checkboxes = divDeItem.querySelectorAll('input[type="checkbox"]');

        checkboxes.forEach(function(checkbox) {
            if (!checkbox.checked) {
                checkbox.nextSibling.remove();
                checkbox.remove();
            }
        });
    }
}
function enviarEmail() {
    var resultado = document.getElementById('resultado').innerText;
    var corpoDoEmail = encodeURIComponent(resultado);
    window.open(`mailto:?subject=Relação Inventario 2023 Preview&body=${corpoDoEmail}`);
}
