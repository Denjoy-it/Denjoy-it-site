"""
Tenant Knowledge Base API — Flask Blueprint
Provides CRUD for: assets, VLANs, KB-pages (Markdown) and contacts.
Data is stored in SQLite per tenant: data/tenants/{tenant_id}/kb.sqlite
"""
import os
import sqlite3
from flask import Blueprint, request, jsonify

kb = Blueprint('kb', __name__)

ROOT = os.path.dirname(os.path.abspath(__file__))

# When running as a packaged desktop app, M365_DATA_DIR points to a writable
# location outside the read-only bundle (e.g. ~/Library/Application Support/M365Tool).
# In development it falls back to ROOT/data.
_DATA_ROOT = os.environ.get('M365_DATA_DIR') or os.path.join(ROOT, 'data')


# ---------------------------------------------------------------------------
# Database helpers
# ---------------------------------------------------------------------------

def _db_path(tenant_id: str) -> str:
    safe = os.path.basename(tenant_id.replace('..', ''))
    path = os.path.join(_DATA_ROOT, 'tenants', safe)
    os.makedirs(path, exist_ok=True)
    return os.path.join(path, 'kb.sqlite')


def _get_conn(tenant_id: str) -> sqlite3.Connection:
    conn = sqlite3.connect(_db_path(tenant_id))
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA foreign_keys=ON')
    return conn


def _init_schema(conn: sqlite3.Connection) -> None:
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS asset_types (
        id   INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        icon TEXT DEFAULT '🖥️'
    );

    INSERT OR IGNORE INTO asset_types (name, icon) VALUES
        ('switch',   '🔀'),
        ('router',   '🌐'),
        ('firewall', '🛡️'),
        ('ap',       '📡'),
        ('server',   '🖥️'),
        ('vlan',     '🏷️'),
        ('subnet',   '🕸️'),
        ('circuit',  '🔌');

    CREATE TABLE IF NOT EXISTS assets (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        asset_type_id INTEGER REFERENCES asset_types(id),
        name          TEXT NOT NULL,
        hostname      TEXT,
        ip_address    TEXT,
        location      TEXT,
        vendor        TEXT,
        model         TEXT,
        firmware      TEXT,
        serial        TEXT,
        notes         TEXT,
        is_active     INTEGER DEFAULT 1,
        created_at    TEXT DEFAULT (datetime('now')),
        updated_at    TEXT DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS vlans (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        vlan_id     INTEGER NOT NULL,
        name        TEXT NOT NULL,
        subnet      TEXT,
        gateway     TEXT,
        description TEXT,
        purpose     TEXT DEFAULT 'user',
        notes       TEXT,
        created_at  TEXT DEFAULT (datetime('now')),
        updated_at  TEXT DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS asset_vlans (
        id       INTEGER PRIMARY KEY AUTOINCREMENT,
        asset_id INTEGER REFERENCES assets(id) ON DELETE CASCADE,
        vlan_id  INTEGER REFERENCES vlans(id)  ON DELETE CASCADE,
        port     TEXT,
        mode     TEXT DEFAULT 'trunk'
    );

    CREATE TABLE IF NOT EXISTS kb_pages (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        title       TEXT NOT NULL,
        content     TEXT DEFAULT '',
        category    TEXT DEFAULT 'network',
        order_index INTEGER DEFAULT 0,
        updated_at  TEXT DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS contacts (
        id                 INTEGER PRIMARY KEY AUTOINCREMENT,
        name               TEXT NOT NULL,
        role               TEXT,
        phone              TEXT,
        email              TEXT,
        is_primary_contact INTEGER DEFAULT 0,
        notes              TEXT,
        created_at         TEXT DEFAULT (datetime('now'))
    );
    """)
    conn.commit()


def _row_to_dict(row) -> dict:
    return dict(row) if row else None


def _rows_to_list(rows) -> list:
    return [dict(r) for r in rows]


def _check_tenant(tenant_id: str):
    if not tenant_id or len(tenant_id) > 64:
        return jsonify({'error': 'Invalid tenant_id'}), 400
    return None


# ---------------------------------------------------------------------------
# Asset types
# ---------------------------------------------------------------------------

@kb.route('/api/kb/<tenant_id>/asset-types', methods=['GET'])
def list_asset_types(tenant_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        rows = conn.execute('SELECT * FROM asset_types ORDER BY name').fetchall()
    return jsonify(_rows_to_list(rows))


# ---------------------------------------------------------------------------
# Assets
# ---------------------------------------------------------------------------

@kb.route('/api/kb/<tenant_id>/assets', methods=['GET'])
def list_assets(tenant_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    asset_type = request.args.get('type')
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        sql = (
            'SELECT a.*, t.name as type_name, t.icon as type_icon '
            'FROM assets a LEFT JOIN asset_types t ON a.asset_type_id = t.id'
        )
        params = []
        if asset_type:
            sql += ' WHERE t.name = ?'
            params.append(asset_type)
        sql += ' ORDER BY a.name'
        rows = conn.execute(sql, params).fetchall()
    return jsonify(_rows_to_list(rows))


@kb.route('/api/kb/<tenant_id>/assets', methods=['POST'])
def create_asset(tenant_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    data = request.get_json() or {}
    if not data.get('name'):
        return jsonify({'error': 'Field "name" is required'}), 400
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        cur = conn.execute(
            'INSERT INTO assets '
            '(asset_type_id, name, hostname, ip_address, location, vendor, model, firmware, serial, notes, is_active) '
            'VALUES (?,?,?,?,?,?,?,?,?,?,?)',
            (
                data.get('asset_type_id'),
                data['name'],
                data.get('hostname'),
                data.get('ip_address'),
                data.get('location'),
                data.get('vendor'),
                data.get('model'),
                data.get('firmware'),
                data.get('serial'),
                data.get('notes'),
                int(data.get('is_active', 1)),
            ),
        )
        conn.commit()
        row = conn.execute(
            'SELECT a.*, t.name as type_name, t.icon as type_icon '
            'FROM assets a LEFT JOIN asset_types t ON a.asset_type_id=t.id WHERE a.id=?',
            (cur.lastrowid,),
        ).fetchone()
    return jsonify(_row_to_dict(row)), 201


@kb.route('/api/kb/<tenant_id>/assets/<int:asset_id>', methods=['GET'])
def get_asset(tenant_id, asset_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        row = conn.execute(
            'SELECT a.*, t.name as type_name, t.icon as type_icon '
            'FROM assets a LEFT JOIN asset_types t ON a.asset_type_id=t.id WHERE a.id=?',
            (asset_id,),
        ).fetchone()
    if not row:
        return jsonify({'error': 'Not found'}), 404
    return jsonify(_row_to_dict(row))


@kb.route('/api/kb/<tenant_id>/assets/<int:asset_id>', methods=['PUT'])
def update_asset(tenant_id, asset_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    data = request.get_json() or {}
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        conn.execute(
            'UPDATE assets SET asset_type_id=?, name=?, hostname=?, ip_address=?, location=?, '
            'vendor=?, model=?, firmware=?, serial=?, notes=?, is_active=?, updated_at=datetime(\'now\') '
            'WHERE id=?',
            (
                data.get('asset_type_id'),
                data.get('name'),
                data.get('hostname'),
                data.get('ip_address'),
                data.get('location'),
                data.get('vendor'),
                data.get('model'),
                data.get('firmware'),
                data.get('serial'),
                data.get('notes'),
                int(data.get('is_active', 1)),
                asset_id,
            ),
        )
        conn.commit()
        row = conn.execute(
            'SELECT a.*, t.name as type_name, t.icon as type_icon '
            'FROM assets a LEFT JOIN asset_types t ON a.asset_type_id=t.id WHERE a.id=?',
            (asset_id,),
        ).fetchone()
    if not row:
        return jsonify({'error': 'Not found'}), 404
    return jsonify(_row_to_dict(row))


@kb.route('/api/kb/<tenant_id>/assets/<int:asset_id>', methods=['DELETE'])
def delete_asset(tenant_id, asset_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        conn.execute('DELETE FROM assets WHERE id=?', (asset_id,))
        conn.commit()
    return jsonify({'ok': True})


# ---------------------------------------------------------------------------
# VLANs
# ---------------------------------------------------------------------------

@kb.route('/api/kb/<tenant_id>/vlans', methods=['GET'])
def list_vlans(tenant_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        rows = conn.execute('SELECT * FROM vlans ORDER BY vlan_id').fetchall()
    return jsonify(_rows_to_list(rows))


@kb.route('/api/kb/<tenant_id>/vlans', methods=['POST'])
def create_vlan(tenant_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    data = request.get_json() or {}
    if not data.get('vlan_id') or not data.get('name'):
        return jsonify({'error': 'vlan_id and name are required'}), 400
    try:
        vlan_num = int(data['vlan_id'])
        if not (1 <= vlan_num <= 4094):
            raise ValueError
    except (ValueError, TypeError):
        return jsonify({'error': 'vlan_id must be 1-4094'}), 400
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        cur = conn.execute(
            'INSERT INTO vlans (vlan_id, name, subnet, gateway, description, purpose, notes) '
            'VALUES (?,?,?,?,?,?,?)',
            (
                vlan_num,
                data['name'],
                data.get('subnet'),
                data.get('gateway'),
                data.get('description'),
                data.get('purpose', 'user'),
                data.get('notes'),
            ),
        )
        conn.commit()
        row = conn.execute('SELECT * FROM vlans WHERE id=?', (cur.lastrowid,)).fetchone()
    return jsonify(_row_to_dict(row)), 201


@kb.route('/api/kb/<tenant_id>/vlans/<int:vlan_db_id>', methods=['PUT'])
def update_vlan(tenant_id, vlan_db_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    data = request.get_json() or {}
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        conn.execute(
            'UPDATE vlans SET vlan_id=?, name=?, subnet=?, gateway=?, description=?, purpose=?, notes=?, '
            'updated_at=datetime(\'now\') WHERE id=?',
            (
                data.get('vlan_id'),
                data.get('name'),
                data.get('subnet'),
                data.get('gateway'),
                data.get('description'),
                data.get('purpose', 'user'),
                data.get('notes'),
                vlan_db_id,
            ),
        )
        conn.commit()
        row = conn.execute('SELECT * FROM vlans WHERE id=?', (vlan_db_id,)).fetchone()
    if not row:
        return jsonify({'error': 'Not found'}), 404
    return jsonify(_row_to_dict(row))


@kb.route('/api/kb/<tenant_id>/vlans/<int:vlan_db_id>', methods=['DELETE'])
def delete_vlan(tenant_id, vlan_db_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        conn.execute('DELETE FROM vlans WHERE id=?', (vlan_db_id,))
        conn.commit()
    return jsonify({'ok': True})


# ---------------------------------------------------------------------------
# KB Pages (Markdown docs)
# ---------------------------------------------------------------------------

@kb.route('/api/kb/<tenant_id>/pages', methods=['GET'])
def list_pages(tenant_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        rows = conn.execute(
            'SELECT id, title, category, order_index, updated_at FROM kb_pages ORDER BY order_index, title'
        ).fetchall()
    return jsonify(_rows_to_list(rows))


@kb.route('/api/kb/<tenant_id>/pages', methods=['POST'])
def create_page(tenant_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    data = request.get_json() or {}
    if not data.get('title'):
        return jsonify({'error': 'title is required'}), 400
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        cur = conn.execute(
            'INSERT INTO kb_pages (title, content, category, order_index) VALUES (?,?,?,?)',
            (data['title'], data.get('content', ''), data.get('category', 'network'), data.get('order_index', 0)),
        )
        conn.commit()
        row = conn.execute('SELECT * FROM kb_pages WHERE id=?', (cur.lastrowid,)).fetchone()
    return jsonify(_row_to_dict(row)), 201


@kb.route('/api/kb/<tenant_id>/pages/<int:page_id>', methods=['GET'])
def get_page(tenant_id, page_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        row = conn.execute('SELECT * FROM kb_pages WHERE id=?', (page_id,)).fetchone()
    if not row:
        return jsonify({'error': 'Not found'}), 404
    return jsonify(_row_to_dict(row))


@kb.route('/api/kb/<tenant_id>/pages/<int:page_id>', methods=['PUT'])
def update_page(tenant_id, page_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    data = request.get_json() or {}
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        conn.execute(
            'UPDATE kb_pages SET title=?, content=?, category=?, order_index=?, updated_at=datetime(\'now\') WHERE id=?',
            (data.get('title'), data.get('content', ''), data.get('category', 'network'), data.get('order_index', 0), page_id),
        )
        conn.commit()
        row = conn.execute('SELECT * FROM kb_pages WHERE id=?', (page_id,)).fetchone()
    if not row:
        return jsonify({'error': 'Not found'}), 404
    return jsonify(_row_to_dict(row))


@kb.route('/api/kb/<tenant_id>/pages/<int:page_id>', methods=['DELETE'])
def delete_page(tenant_id, page_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        conn.execute('DELETE FROM kb_pages WHERE id=?', (page_id,))
        conn.commit()
    return jsonify({'ok': True})


# ---------------------------------------------------------------------------
# Contacts
# ---------------------------------------------------------------------------

@kb.route('/api/kb/<tenant_id>/contacts', methods=['GET'])
def list_contacts(tenant_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        rows = conn.execute(
            'SELECT * FROM contacts ORDER BY is_primary_contact DESC, name'
        ).fetchall()
    return jsonify(_rows_to_list(rows))


@kb.route('/api/kb/<tenant_id>/contacts', methods=['POST'])
def create_contact(tenant_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    data = request.get_json() or {}
    if not data.get('name'):
        return jsonify({'error': 'name is required'}), 400
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        cur = conn.execute(
            'INSERT INTO contacts (name, role, phone, email, is_primary_contact, notes) VALUES (?,?,?,?,?,?)',
            (
                data['name'],
                data.get('role'),
                data.get('phone'),
                data.get('email'),
                int(data.get('is_primary_contact', 0)),
                data.get('notes'),
            ),
        )
        conn.commit()
        row = conn.execute('SELECT * FROM contacts WHERE id=?', (cur.lastrowid,)).fetchone()
    return jsonify(_row_to_dict(row)), 201


@kb.route('/api/kb/<tenant_id>/contacts/<int:contact_id>', methods=['PUT'])
def update_contact(tenant_id, contact_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    data = request.get_json() or {}
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        conn.execute(
            'UPDATE contacts SET name=?, role=?, phone=?, email=?, is_primary_contact=?, notes=? WHERE id=?',
            (
                data.get('name'),
                data.get('role'),
                data.get('phone'),
                data.get('email'),
                int(data.get('is_primary_contact', 0)),
                data.get('notes'),
                contact_id,
            ),
        )
        conn.commit()
        row = conn.execute('SELECT * FROM contacts WHERE id=?', (contact_id,)).fetchone()
    if not row:
        return jsonify({'error': 'Not found'}), 404
    return jsonify(_row_to_dict(row))


@kb.route('/api/kb/<tenant_id>/contacts/<int:contact_id>', methods=['DELETE'])
def delete_contact(tenant_id, contact_id):
    err = _check_tenant(tenant_id)
    if err:
        return err
    with _get_conn(tenant_id) as conn:
        _init_schema(conn)
        conn.execute('DELETE FROM contacts WHERE id=?', (contact_id,))
        conn.commit()
    return jsonify({'ok': True})
