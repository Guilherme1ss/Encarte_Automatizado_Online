import sqlite3
import json
from pathlib import Path

DB_FILE = "./data/produtos.db"
JSON_FILE = "./data/default_url.json"

def criar_tabelas(conn):
    cur = conn.cursor()

    cur.execute("PRAGMA foreign_keys = ON;")

    cur.execute("""
        CREATE TABLE IF NOT EXISTS produtos (
            id TEXT PRIMARY KEY,
            posicao INTEGER,
            nome TEXT,
            url TEXT
        );
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS produto_codigos (
            produto_id TEXT,
            codigo TEXT,
            PRIMARY KEY (produto_id, codigo),
            FOREIGN KEY (produto_id) REFERENCES produtos(id) ON DELETE CASCADE
        );
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS produto_eans (
            produto_id TEXT,
            ean TEXT,
            PRIMARY KEY (produto_id, ean),
            FOREIGN KEY (produto_id) REFERENCES produtos(id) ON DELETE CASCADE
        );
    """)

    conn.commit()


def importar_json(conn):
    with open(JSON_FILE, "r", encoding="utf-8") as f:
        produtos = json.load(f)

    cur = conn.cursor()

    for p in produtos:
        pid = p["id"]

        # salvar produto
        cur.execute("""
            INSERT OR REPLACE INTO produtos (id, posicao, nome, url)
            VALUES (?, ?, ?, ?)
        """, (pid, p["posição"], p["nome"], p["url"]))

        # salvar códigos
        for c in p.get("codigo", []):
            cur.execute("""
                INSERT OR IGNORE INTO produto_codigos (produto_id, codigo)
                VALUES (?, ?)
            """, (pid, c))

        # salvar EANs
        for e in p.get("eans", []):
            cur.execute("""
                INSERT OR IGNORE INTO produto_eans (produto_id, ean)
                VALUES (?, ?)
            """, (pid, e))

    conn.commit()


def main():
    Path("data").mkdir(exist_ok=True)

    conn = sqlite3.connect(DB_FILE)
    criar_tabelas(conn)
    importar_json(conn)
    conn.close()

    print("Banco criado e dados importados com sucesso!")


if __name__ == "__main__":
    main()
