"""Gerador reproduzível de dados fictícios para um minimercado.

Executar na raiz do repositório:
    python python/gerar_dados.py

Se desejar substituir explicitamente os três CSVs de desenvolvimento:
    python python/gerar_dados.py --sobrescrever

ATENÇÃO: não regenere a base após começar a registrar movimentações reais
na pasta de trabalho do Excel. Os CSVs constituem o histórico inicial.

Os dados são simulados; custos são referências fictícias, não cotações reais.
As inconsistências de texto são deliberadas para tratamento no Power Query.
"""

import argparse
import calendar
import csv
import random
from collections import Counter
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

SEMENTE = 20261123
QTD_PRODUTOS = 150
QTD_MOVIMENTACOES = 6000
DATA_INICIAL = date(2026, 1, 1)
DATA_FINAL = date(2026, 10, 31)
PASTA_DADOS = Path(__file__).resolve().parent.parent / "dados"

# Mantemos exatamente as colunas previstas na modelagem do projeto.
COLUNAS_PRODUTOS = [
    "id_produto", "nome_produto", "categoria", "id_fornecedor",
    "estoque_inicial", "estoque_minimo", "custo_unitario",
]
COLUNAS_FORNECEDORES = ["id_fornecedor", "nome_fornecedor", "cidade", "estado"]
COLUNAS_MOVIMENTACOES = [
    "id_movimento", "data_movimento", "id_produto", "tipo", "quantidade",
]

# Valores nominais fictícios de custo de aquisição, em R$.
# Cada lista segue a ordem das cinco apresentações da respectiva família.
# Pequenas variações de até aproximadamente 8% serão aplicadas na geração.
CUSTOS_REFERENCIA = {
    "Arroz Branco": [5.20, 10.20, 23.40, 6.60, 5.60],
    "Feijão": [6.20, 7.40, 4.80, 5.10, 7.00],
    "Açúcar": [3.90, 7.40, 4.80, 5.20, 6.80],
    "Macarrão": [3.70, 4.10, 4.50, 5.90, 7.20],
    "Farinha": [4.20, 6.30, 3.30, 5.10, 4.70],
    "Biscoito": [3.80, 4.20, 4.60, 3.30, 4.00],
    "Leite": [4.80, 5.20, 4.90, 6.10, 15.90],
    "Suco": [8.40, 7.20, 8.80, 7.00, 6.90],
    "Água": [1.80, 3.40, 6.80, 2.90, 4.20],
    "Refrigerante": [7.60, 7.20, 7.90, 7.80, 3.40],
    "Café": [16.50, 30.80, 18.40, 34.40, 12.80],
    "Óleo": [7.30, 11.20, 12.10, 10.90, 16.00],
    "Molho": [2.80, 4.20, 5.50, 7.10, 6.20],
    "Milho Verde": [3.40, 4.50, 3.10, 4.40, 3.90],
    "Ervilha": [3.60, 4.60, 3.40, 4.80, 4.20],
    "Detergente": [2.50, 2.60, 2.70, 2.80, 4.60],
    "Sabão em Pó": [5.70, 8.30, 10.50, 16.20, 20.50],
    "Desinfetante": [3.20, 3.10, 3.30, 8.50, 8.10],
    "Água Sanitária": [2.10, 3.50, 5.90, 8.10, 12.40],
    "Amaciante": [4.10, 11.40, 4.40, 11.90, 12.20],
    "Sabonete": [1.80, 2.10, 2.10, 2.20, 9.90],
    "Shampoo": [10.80, 12.20, 14.80, 9.30, 16.80],
    "Condicionador": [11.20, 12.10, 15.80, 10.90, 16.20],
    "Creme Dental": [4.10, 5.50, 5.80, 4.00, 7.80],
    "Papel Higiênico": [6.10, 8.30, 10.90, 14.10, 18.90],
    "Iogurte": [3.00, 3.40, 3.40, 6.60, 7.00],
    "Queijo": [9.50, 18.40, 9.90, 20.00, 5.50],
    "Manteiga": [11.00, 11.50, 24.00, 25.00, 13.60],
    "Legumes Congelados": [7.40, 11.10, 8.70, 7.60, 7.30],
    "Polpa de Fruta": [2.50, 2.60, 2.30, 2.60, 3.00],
}

# Fornecedor principal por categoria; mesmo fornecedor pode atender mais
# de uma categoria, como acontece em distribuidores generalistas.
FORNECEDORES_POR_CATEGORIA = {
    "Alimentos": ["F001", "F003", "F005", "F009"],
    "Bebidas": ["F002", "F003", "F006"],
    "Mercearia": ["F001", "F007", "F008"],
    "Limpeza": ["F004", "F006", "F011"],
    "Higiene": ["F004", "F008", "F010"],
    "Laticínios": ["F005", "F009", "F012"],
    "Congelados": ["F007", "F010", "F012"],
}

# Peso aproximado da procura de famílias de produtos.
PESO_DEMANDA_FAMILIA = {
    "Arroz Branco": 3.3, "Feijão": 2.8, "Açúcar": 2.4,
    "Leite": 3.2, "Água": 3.3, "Café": 2.9,
    "Refrigerante": 2.7, "Macarrão": 2.5, "Detergente": 2.2,
    "Papel Higiênico": 2.1, "Sabonete": 1.9, "Biscoito": 2.3,
    "Óleo": 2.3, "Iogurte": 2.0, "Queijo": 1.6,
    "Legumes Congelados": 1.0, "Polpa de Fruta": 1.1,
}

# Mais saídas no fim do período simulam maior demanda no último trimestre.
PESOS_MESES = [0.88, 0.90, 0.95, 0.98, 1.00, 1.02, 1.03, 1.07, 1.12, 1.25]

# Erros NÃO críticos deliberados: nunca afetam códigos, datas ou quantidades.
CATEGORIAS_COM_VARIACAO = {
    6: "ALIMENTOS", 14: "alimentos", 21: "AlIMenTos",
    32: "bebidas", 43: "BEBIDAS", 49: "BeBiDaS",
    77: "mercearia", 87: "LIMPEZA", 99: "limpeza",
    108: "higiene", 114: "HIGIENE", 123: "HiGiEnE",
}
NOMES_COM_BORDAS = {7, 29, 61, 103, 141, 16, 120}
NOMES_COM_ESPACO_DUPLO = {18, 58, 94, 133}

FAMILIAS = [
    ("Alimentos", "Arroz Branco", ["1 kg", "2 kg", "5 kg", "Integral 1 kg", "Parboilizado 1 kg"]),
    ("Alimentos", "Feijão", ["Carioca 1 kg", "Preto 1 kg", "Fradinho 500 g", "Branco 500 g", "Vermelho 1 kg"]),
    ("Alimentos", "Açúcar", ["Cristal 1 kg", "Cristal 2 kg", "Refinado 1 kg", "Demerara 500 g", "Mascavo 500 g"]),
    ("Alimentos", "Macarrão", ["Espaguete 500 g", "Parafuso 500 g", "Penne 500 g", "Integral 500 g", "Lasanha 500 g"]),
    ("Alimentos", "Farinha", ["de Trigo 1 kg", "de Mandioca 1 kg", "de Milho 500 g", "de Aveia 500 g", "de Rosca 500 g"]),
    ("Alimentos", "Biscoito", ["Água e Sal 200 g", "Maisena 200 g", "Integral 150 g", "Recheado Chocolate 120 g", "Cream Cracker 200 g"]),
    ("Bebidas", "Leite", ["Integral 1 L", "Desnatado 1 L", "Semidesnatado 1 L", "Zero Lactose 1 L", "Em Pó 400 g"]),
    ("Bebidas", "Suco", ["de Uva 1 L", "de Laranja 1 L", "de Maçã 1 L", "de Caju 1 L", "de Goiaba 1 L"]),
    ("Bebidas", "Água", ["Mineral 500 ml", "Mineral 1,5 L", "Mineral 5 L", "com Gás 500 ml", "com Gás 1 L"]),
    ("Bebidas", "Refrigerante", ["Cola 2 L", "Guaraná 2 L", "Laranja 2 L", "Limão 2 L", "Cola 350 ml"]),
    ("Bebidas", "Café", ["Tradicional 250 g", "Tradicional 500 g", "Extra Forte 250 g", "Extra Forte 500 g", "Solúvel 100 g"]),
    ("Mercearia", "Óleo", ["de Soja 900 ml", "de Girassol 900 ml", "de Canola 900 ml", "de Milho 900 ml", "de Coco 200 ml"]),
    ("Mercearia", "Molho", ["de Tomate 300 g", "de Tomate 500 g", "de Pimenta 150 ml", "Shoyu 150 ml", "Inglês 150 ml"]),
    ("Mercearia", "Milho Verde", ["Lata 170 g", "Lata 280 g", "Sachê 170 g", "Sachê 280 g", "Conserva 200 g"]),
    ("Mercearia", "Ervilha", ["Lata 170 g", "Lata 280 g", "Sachê 170 g", "Sachê 280 g", "Conserva 200 g"]),
    ("Limpeza", "Detergente", ["Neutro 500 ml", "Limão 500 ml", "Coco 500 ml", "Maçã 500 ml", "Neutro 1 L"]),
    ("Limpeza", "Sabão em Pó", ["500 g", "800 g", "1 kg", "1,6 kg", "2 kg"]),
    ("Limpeza", "Desinfetante", ["Lavanda 500 ml", "Pinho 500 ml", "Eucalipto 500 ml", "Lavanda 2 L", "Pinho 2 L"]),
    ("Limpeza", "Água Sanitária", ["500 ml", "1 L", "2 L", "3 L", "5 L"]),
    ("Limpeza", "Amaciante", ["Flores 500 ml", "Flores 2 L", "Lavanda 500 ml", "Lavanda 2 L", "Neutro 2 L"]),
    ("Higiene", "Sabonete", ["Neutro 85 g", "Erva-Doce 85 g", "Lavanda 85 g", "Glicerina 90 g", "Líquido 250 ml"]),
    ("Higiene", "Shampoo", ["Neutro 200 ml", "Hidratação 200 ml", "Anticaspa 200 ml", "Infantil 200 ml", "Neutro 350 ml"]),
    ("Higiene", "Condicionador", ["Neutro 200 ml", "Hidratação 200 ml", "Cachos 200 ml", "Infantil 200 ml", "Neutro 350 ml"]),
    ("Higiene", "Creme Dental", ["Menta 70 g", "Menta 90 g", "Branqueador 70 g", "Infantil 50 g", "Sensibilidade 90 g"]),
    ("Higiene", "Papel Higiênico", ["4 Rolos 20 m", "4 Rolos 30 m", "8 Rolos 20 m", "8 Rolos 30 m", "12 Rolos 30 m"]),
    ("Laticínios", "Iogurte", ["Natural 170 g", "Morango 170 g", "Coco 170 g", "Natural 500 g", "Morango 500 g"]),
    ("Laticínios", "Queijo", ["Muçarela 150 g", "Muçarela 300 g", "Prato 150 g", "Minas 300 g", "Parmesão Ralado 50 g"]),
    ("Laticínios", "Manteiga", ["com Sal 200 g", "sem Sal 200 g", "com Sal 500 g", "sem Sal 500 g", "Extra 200 g"]),
    ("Congelados", "Legumes Congelados", ["Mix 300 g", "Mix 500 g", "Brócolis 300 g", "Ervilha 300 g", "Cenoura 300 g"]),
    ("Congelados", "Polpa de Fruta", ["Acerola 100 g", "Caju 100 g", "Goiaba 100 g", "Manga 100 g", "Maracujá 100 g"]),
]

FORNECEDORES = [
    ("F001", "Distribuidora Agreste", "Caruaru", "PE"),
    ("F002", "Atacado Sertão", "Petrolina", "PE"),
    ("F003", "Comercial Nordeste", "Recife", "PE"),
    ("F004", "Logística Boa Feira", "Garanhuns", "PE"),
    ("F005", "Central de Alimentos", "Santa Cruz do Capibaribe", "PE"),
    ("F006", "Distribuidora Boa Compra", "Toritama", "PE"),
    ("F007", "Rede Abastecer", "Belo Jardim", "PE"),
    ("F008", "Comercial São José", "Bezerros", "PE"),
    ("F009", "Atacado Esperança", "Vitória de Santo Antão", "PE"),
    ("F010", "Distribuidora Litoral", "Olinda", "PE"),
    ("F011", "Centro de Suprimentos", "Gravatá", "PE"),
    ("F012", "Atacado Vale Verde", "Arcoverde", "PE"),
]


def custo_formatado(referencia, aleatorio):
    """Retorna custo fictício com vírgula decimal para Excel pt-BR."""
    variacao = Decimal(str(aleatorio.uniform(0.93, 1.07)))
    custo = (Decimal(str(referencia)) * variacao).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP
    )
    return f"{custo:.2f}".replace(".", ",")


def gerar_fornecedores():
    fornecedores = []
    for codigo, nome, cidade, estado in FORNECEDORES:
        # Inconsistências intencionais de apresentação, não de identidade.
        if codigo == "F002":
            cidade = cidade.lower()
        elif codigo == "F004":
            cidade = cidade.upper()
        elif codigo == "F007":
            nome = f"  {nome}  "
        elif codigo == "F009":
            estado = estado.lower()
        elif codigo == "F012":
            cidade = ""  # Campo opcional ausente: não altera o vínculo pelo ID.
        fornecedores.append(dict(zip(COLUNAS_FORNECEDORES, (
            codigo, nome, cidade, estado
        ))))
    return fornecedores


def gerar_produtos(aleatorio):
    produtos = []
    for indice_familia, (categoria, familia, apresentacoes) in enumerate(FAMILIAS):
        fornecedores_permitidos = FORNECEDORES_POR_CATEGORIA[categoria]
        for indice_apresentacao, apresentacao in enumerate(apresentacoes):
            numero = len(produtos) + 1
            nome = f"{familia} {apresentacao}"
            if numero in NOMES_COM_ESPACO_DUPLO:
                nome = nome.replace(" ", "  ", 1)
            if numero in NOMES_COM_BORDAS:
                nome = f"  {nome}  "

            estoque_minimo = aleatorio.randint(12, 28)
            estoque_inicial = estoque_minimo + aleatorio.randint(22, 70)
            fornecedor = fornecedores_permitidos[
                (indice_familia + indice_apresentacao) % len(fornecedores_permitidos)
            ]
            produtos.append({
                "id_produto": f"P{numero:03d}",
                "nome_produto": nome,
                "categoria": CATEGORIAS_COM_VARIACAO.get(numero, categoria),
                "id_fornecedor": fornecedor,
                "estoque_inicial": estoque_inicial,
                "estoque_minimo": estoque_minimo,
                "custo_unitario": custo_formatado(
                    CUSTOS_REFERENCIA[familia][indice_apresentacao], aleatorio
                ),
            })
    return produtos


def gerar_datas(aleatorio):
    meses = aleatorio.choices(range(1, 11), weights=PESOS_MESES,
                              k=QTD_MOVIMENTACOES)
    datas = []
    for mes in meses:
        ultimo_dia = calendar.monthrange(2026, mes)[1]
        datas.append(date(2026, mes, aleatorio.randint(1, ultimo_dia)))
    return sorted(datas)


def gerar_movimentacoes(aleatorio, produtos):
    saldos = {p["id_produto"]: p["estoque_inicial"] for p in produtos}
    pesos = []
    for p in produtos:
        familia = next(f for _, f, apresentacoes in FAMILIAS
                       if p["nome_produto"].strip().replace("  ", " ").startswith(f + " "))
        pesos.append(PESO_DEMANDA_FAMILIA.get(familia, 1.45))

    movimentos = []
    for indice, data_movimento in enumerate(gerar_datas(aleatorio), 1):
        produto = aleatorio.choices(produtos, weights=pesos, k=1)[0]
        codigo = produto["id_produto"]
        saldo = saldos[codigo]
        minimo = produto["estoque_minimo"]

        # No fim de outubro, algumas reposições demoram mais a chegar.
        # Isso cria alertas de estoque mínimo sem falsear o saldo.
        atraso_reposicao = data_movimento >= date(2026, 10, 15)
        prob_repor_baixo = 0.52 if atraso_reposicao else 0.78
        if saldo == 0 or (saldo <= minimo and aleatorio.random() < prob_repor_baixo):
            entrada = True
        else:
            entrada = aleatorio.random() < 0.025

        if entrada:
            tipo = "ENTRADA"
            # Reposições maiores do que vendas unitárias, sem exagerar o saldo.
            quantidade = aleatorio.randint(18, 39)
            if codigo in {"P001", "P002", "P031", "P032"}:
                quantidade += aleatorio.randint(4, 16)
            saldos[codigo] += quantidade
        else:
            tipo = "SAIDA"
            quantidade = aleatorio.randint(1, min(11, saldo))
            saldos[codigo] -= quantidade

        movimentos.append({
            "id_movimento": f"H{indice:05d}",
            "data_movimento": data_movimento.strftime("%d/%m/%Y"),
            "id_produto": codigo,
            "tipo": tipo,
            "quantidade": quantidade,
        })
    return movimentos


def exigir(condicao, mensagem):
    if not condicao:
        raise ValueError(mensagem)


def validar(produtos, fornecedores, movimentos):
    """Confere a lógica do gerador; limpeza de texto fica no Power Query."""
    exigir(len(produtos) == QTD_PRODUTOS, "Quantidade incorreta de produtos")
    exigir(len(fornecedores) == 12, "Quantidade incorreta de fornecedores")
    exigir(len(movimentos) == QTD_MOVIMENTACOES,
           "Quantidade incorreta de movimentações")

    ids_produto = [p["id_produto"] for p in produtos]
    ids_fornecedor = [f["id_fornecedor"] for f in fornecedores]
    ids_movimento = [m["id_movimento"] for m in movimentos]
    for descricao, ids in [
        ("produto", ids_produto),
        ("fornecedor", ids_fornecedor),
        ("movimentação", ids_movimento),
    ]:
        exigir(len(ids) == len(set(ids)), f"ID de {descricao} duplicado")

    conjunto_produtos = set(ids_produto)
    conjunto_fornecedores = set(ids_fornecedor)
    exigir(all(p["id_fornecedor"] in conjunto_fornecedores for p in produtos),
           "Produto com fornecedor inexistente")
    exigir(all(m["id_produto"] in conjunto_produtos for m in movimentos),
           "Movimentação com produto inexistente")

    categorias_validas = {categoria.casefold() for categoria, _, _ in FAMILIAS}
    saldos = {}
    minimos = {}
    for p in produtos:
        exigir(p["categoria"].strip().casefold() in categorias_validas,
               "Categoria não reconhecida")
        exigir(isinstance(p["estoque_inicial"], int)
               and p["estoque_inicial"] >= 0, "Estoque inicial inválido")
        exigir(isinstance(p["estoque_minimo"], int)
               and p["estoque_minimo"] > 0, "Estoque mínimo inválido")
        exigir(Decimal(p["custo_unitario"].replace(",", ".")) > 0,
               "Custo inválido")
        saldos[p["id_produto"]] = p["estoque_inicial"]
        minimos[p["id_produto"]] = p["estoque_minimo"]

    data_anterior = DATA_INICIAL
    for m in movimentos:
        data_movimento = datetime.strptime(m["data_movimento"], "%d/%m/%Y").date()
        exigir(DATA_INICIAL <= data_movimento <= DATA_FINAL,
               "Data fora do período")
        exigir(data_movimento >= data_anterior,
               "Movimentações não estão em ordem cronológica")
        data_anterior = data_movimento
        exigir(m["tipo"] in {"ENTRADA", "SAIDA"}, "Tipo inválido")
        exigir(isinstance(m["quantidade"], int) and m["quantidade"] > 0,
               "Quantidade não positiva")
        codigo = m["id_produto"]
        sinal = 1 if m["tipo"] == "ENTRADA" else -1
        saldos[codigo] += sinal * m["quantidade"]
        exigir(saldos[codigo] >= 0,
               f"Estoque negativo para {codigo} em {m['id_movimento']}")
    exigir(any(saldos[c] <= minimos[c] for c in saldos),
           "A simulação não gerou alertas de estoque")
    return saldos


def gravar_csv(caminho, colunas, linhas):
    # UTF-8 com BOM e ponto e vírgula: adequado ao Excel em português do Brasil.
    with caminho.open("w", encoding="utf-8-sig", newline="") as arquivo:
        escritor = csv.DictWriter(arquivo, fieldnames=colunas, delimiter=";")
        escritor.writeheader()
        escritor.writerows(linhas)


def main():
    parser = argparse.ArgumentParser(
        description="Gera dados sintéticos para o controle de estoque"
    )
    parser.add_argument(
        "--sobrescrever", action="store_true",
        help="Permite recriar a base histórica e substituir os CSVs existentes",
    )
    args = parser.parse_args()
    destinos = {
        "produtos": PASTA_DADOS / "produtos.csv",
        "fornecedores": PASTA_DADOS / "fornecedores.csv",
        "movimentacoes": PASTA_DADOS / "movimentacoes_historicas.csv",
    }
    existentes = [str(caminho) for caminho in destinos.values() if caminho.exists()]
    if existentes and not args.sobrescrever:
        parser.error(
            "Os CSVs já existem e não foram alterados. Para recriar "+
            "EXPLICITAMENTE o histórico, use --sobrescrever: " + ", ".join(existentes)
        )

    aleatorio = random.Random(SEMENTE)
    fornecedores = gerar_fornecedores()
    produtos = gerar_produtos(aleatorio)
    movimentos = gerar_movimentacoes(aleatorio, produtos)
    saldos = validar(produtos, fornecedores, movimentos)

    PASTA_DADOS.mkdir(parents=True, exist_ok=True)
    gravar_csv(destinos["produtos"], COLUNAS_PRODUTOS, produtos)
    gravar_csv(destinos["fornecedores"], COLUNAS_FORNECEDORES, fornecedores)
    gravar_csv(destinos["movimentacoes"], COLUNAS_MOVIMENTACOES, movimentos)

    contagem = Counter(m["tipo"] for m in movimentos)
    minimos = {p["id_produto"]: p["estoque_minimo"] for p in produtos}
    alertas = sum(saldo <= minimos[codigo] for codigo, saldo in saldos.items())
    print("Base histórica criada e validada com sucesso.")
    print(f"Produtos: {len(produtos)} | Fornecedores: {len(fornecedores)}")
    print(f"Movimentações: {len(movimentos)} "
          f"(entradas: {contagem['ENTRADA']}; saídas: {contagem['SAIDA']})")
    print(f"Período: {DATA_INICIAL:%d/%m/%Y} a {DATA_FINAL:%d/%m/%Y}")
    print(f"Estoque final: mínimo {min(saldos.values())}, "
          f"máximo {max(saldos.values())}")
    print(f"Produtos em alerta (saldo <= mínimo): {alertas}")
    print(f"Arquivos salvos em: {PASTA_DADOS}")


if __name__ == "__main__":
    main()
