from flask import Flask, jsonify, request
import mysql.connector
import os

app = Flask(__name__)

def conectar_mysql():
    return mysql.connector.connect(
        host=os.environ.get('DB_HOST'),
        user=os.environ.get('DB_USER'),
        password=os.environ.get('DB_PASSWORD'),
        database=os.environ.get('DB_NAME')
    )

@app.route('/inserir', methods=['POST'])
def inserir_dinamico():
    dados = request.json
    tabela = dados.get("tabela")
    campos = dados.get("campos")
    valores = dados.get("valores")

    if not tabela or not campos or not valores or len(campos) != len(valores):
        return jsonify({'erro': 'Campos inválidos ou incompletos'}), 400

    campos_str = ", ".join(campos)
    placeholders = ", ".join(["%s"] * len(valores))

    try:
        conexao = conectar_mysql()
        cursor = conexao.cursor()

        query = f"INSERT INTO {tabela} ({campos_str}) VALUES ({placeholders})"
        cursor.execute(query, valores)

        conexao.commit()
        return jsonify({'mensagem': '✅ Registro inserido com sucesso!'})

    except mysql.connector.Error as erro:
        return jsonify({'erro': str(erro)}), 500

    finally:
        if 'cursor' in locals(): cursor.close()
        if 'conexao' in locals() and conexao.is_connected(): conexao.close()

@app.route('/atualizar', methods=['PUT'])
def atualizar_dinamico():
    dados = request.json
    tabela = dados.get("tabela")
    campos = dados.get("campos")
    valores = dados.get("valores")
    filtro = dados.get("filtro")  # Ex: {"ID": 5}

    if not tabela or not campos or not valores or not filtro:
        return jsonify({'erro': 'Campos ou filtro incompletos'}), 400

    set_str = ", ".join([f"{campo} = %s" for campo in campos])
    filtro_str = " AND ".join([f"{k} = %s" for k in filtro])
    parametros = valores + list(filtro.values())

    try:
        conexao = conectar_mysql()
        cursor = conexao.cursor()

        query = f"UPDATE {tabela} SET {set_str} WHERE {filtro_str}"
        cursor.execute(query, parametros)

        conexao.commit()

        if cursor.rowcount == 0:
            return jsonify({'mensagem': '⚠️ Nenhum registro atualizado. Filtro pode não ter correspondência.'}), 404
        return jsonify({'mensagem': '🔄 Registro atualizado com sucesso!'})

    except mysql.connector.Error as erro:
        return jsonify({'erro': str(erro)}), 500

    finally:
        if 'cursor' in locals(): cursor.close()
        if 'conexao' in locals() and conexao.is_connected(): conexao.close()

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
