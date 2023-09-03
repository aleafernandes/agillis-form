// Função para salvar os dados no localStorage
function salvarDados() {
    var vendedor = document.getElementById('vendedor').value;
    var nome = document.getElementById('nome').value;
    var email = document.getElementById('email').value;
    var endereco = document.getElementById('endereco').value;
    var bairro = document.getElementById('bairro').value;
    var telefone = document.getElementById('telefone').value;
    var responsavel = document.getElementById('responsavel').value;
    var portaria = document.getElementById('portaria').value;
    var portariaEmpresa = document.getElementById('portaria-e').value;
    var limpezaconservacao = document.getElementById('limpezaconservacao').value;
    var limpezaconservacaoEmpresa = document.getElementById('limpezaconservacao-e').value;
    var vigilanciapatrimonial = document.getElementById('vigilanciapatrimonial').value;
    var vigilanciapatrimonialEmpresa = document.getElementById('vigilanciapatrimonial-e').value;
    var data = document.getElementById('data').value;
    var ocorrencia = document.getElementById('ocorrencia').value;


    var novoDado = {
        vendedor: vendedor,
        nome: nome,
        email: email,
        endereco: endereco,
        bairro: bairro,
        telefone: telefone,
        responsavel: responsavel,
        portaria:  portaria,
        portariaEmpresa:  portariaEmpresa,
        limpezaconservacao:  limpezaconservacao,
        limpezaconservacaoEmpresa:  limpezaconservacaoEmpresa,
        vigilanciapatrimonial:  vigilanciapatrimonial,
        vigilanciapatrimonialEmpresa:  vigilanciapatrimonialEmpresa,
        data: data,
        ocorrencia: ocorrencia
    };

    var dadosExistentes = JSON.parse(localStorage.getItem('dadosFormulario') || '[]');
    dadosExistentes.push(novoDado);

    localStorage.setItem('dadosFormulario', JSON.stringify(dadosExistentes));

    // Limpar o formulário
    document.getElementById('meuFormulario').reset();
}

// Função para gerar o arquivo Excel
function gerarExcel() {
    var dadosExistentes = JSON.parse(localStorage.getItem('dadosFormulario') || '[]');

    if (dadosExistentes.length === 0) {
        alert('Nenhum dado para gerar Excel.');
        return;
    }

    var wb = XLSX.utils.book_new();

    var data = [["Vendedor","Cliente", "Email", "Endereço", "Bairro", "Telefone", "Responsável", "Portaria", "Empresa de Portaria", "Limpeza e Conservação", "Empresa de Limpeza e Conservação", "Vigilância Patrimonial", "Empresa de Vigilância Patrimonial", "Data", "Ocorrência"]];

    for (var i = 0; i < dadosExistentes.length; i++) {
        data.push([dadosExistentes[i].vendedor,dadosExistentes[i].nome, dadosExistentes[i].email, dadosExistentes[i].endereco, dadosExistentes[i].bairro, dadosExistentes[i].telefone, dadosExistentes[i].responsavel, dadosExistentes[i].portaria, dadosExistentes[i].portariaEmpresa, dadosExistentes[i].limpezaconservacao, dadosExistentes[i].limpezaconservacaoEmpresa, dadosExistentes[i].vigilanciapatrimonial, dadosExistentes[i].vigilanciapatrimonialEmpresa, dadosExistentes[i].data, dadosExistentes[i].ocorrencia]);
    }

    var ws = XLSX.utils.aoa_to_sheet(data);

    XLSX.utils.book_append_sheet(wb, ws, "Dados");

    XLSX.writeFile(wb, "dados_formulario.xlsx");
}

// Função para limpar o localStorage com confirmação
function limparLocalStorage() {
    var confirmacao = confirm("Você tem certeza que deseja limpar todos os dados armazenados? Isso apagará todo e qualquer formulário preenchido anteriormente.");

    if (confirmacao) {
        localStorage.removeItem('dadosFormulario');
        alert('Dados limpos com sucesso.');
    }
}
