from quatro import sql_query, scalar_data, tabular_data


# Pull 'prt_no' record from 'planning_lot_quantity' table based on 'plq_id' record
def plq_id_prt_no(config, plq_id):
    sql_exp = f'SELECT trim(prt_no) FROM planning_lot_quantity WHERE plq_id = {plq_id}'
    result_set = sql_query(sql_exp, config.sigm_db_cursor)
    prt_no = scalar_data(result_set)
    return prt_no


# Pull 'prt_desc1' record from 'part' table based on 'prt_no' record
def prt_no_prt_desc(config, prt_no):
    sql_exp = f'SELECT trim(prt_desc2) FROM part WHERE prt_no = \'{prt_no}\''
    result_set = sql_query(sql_exp, config.sigm_db_cursor)
    prt_desc1 = scalar_data(result_set)
    return prt_desc1


# Pull 'plq_qty_per' record from 'planning_lot_quantity' table based on 'plq_id' record
def plq_id_plq_qty_per(config, plq_id):
    sql_exp = f'SELECT plq_qty_per FROM planning_lot_quantity WHERE plq_id = {plq_id}'
    result_set = sql_query(sql_exp, config.sigm_db_cursor)
    plq_qty_per = scalar_data(result_set)
    return plq_qty_per


# Pull 'plq_qty_per' record from 'planning_lot_quantity' table based on 'plq_id' record
def plq_id_plq_note(config, plq_id):
    sql_exp = f'SELECT trim(plq_note) FROM planning_lot_quantity WHERE plq_id = {plq_id}'
    result_set = sql_query(sql_exp, config.sigm_db_cursor)
    plq_note = scalar_data(result_set)
    return plq_note


def orl_id_plq_note(config, orl_id):
    sql_exp = f'SELECT trim(plq_note) FROM planning_lot_quantity WHERE orl_id = {orl_id}'
    result_set = sql_query(sql_exp, config.sigm_db_cursor)
    plq_note = scalar_data(result_set)
    return plq_note


# Pull 'orl_id' record from 'order_line' table based on 'ord_no' record
def ord_no_orl_id(config, ord_no):
    sql_exp = f'SELECT orl_id FROM order_line WHERE ord_no = {ord_no} AND prt_no <> \'\''
    result_set = sql_query(sql_exp, config.sigm_db_cursor)
    orl_ids = tabular_data(result_set)
    return orl_ids


# Pull 'orl_quantity' record from 'order_line' table based on 'orl_id' record
def orl_id_orl_qty(config, orl_id):
    sql_exp = f'SELECT (orl_quantity)::INT FROM order_line WHERE orl_id = {orl_id}'
    result_set = sql_query(sql_exp, config.sigm_db_cursor)
    orl_quantity = scalar_data(result_set)
    return orl_quantity


def orl_id_orl_qty_ready(config, orl_id):
    sql_exp = f'SELECT (orl_qty_ready)::INT FROM order_line WHERE orl_id = {orl_id}'
    result_set = sql_query(sql_exp, config.sigm_db_cursor)
    orl_qty_ready = scalar_data(result_set)
    return orl_qty_ready


# Pull 'prt_no' record from 'order_line' table based on 'orl_id' record
def orl_id_prt_no(config, orl_id):
    sql_exp = f'SELECT trim(prt_no) FROM order_line WHERE orl_id = {orl_id}'
    result_set = sql_query(sql_exp, config.sigm_db_cursor)
    prt_no = scalar_data(result_set)
    return prt_no


# Pull 'prt_no' record from 'order_line' table based on 'orl_id' record
def orl_id_prt_desc(config, orl_id):
    sql_exp = f'SELECT trim(prt_desc) FROM order_line WHERE orl_id = {orl_id}'
    result_set = sql_query(sql_exp, config.sigm_db_cursor)
    prt_desc = scalar_data(result_set)
    return prt_desc


def ord_no_cli_no(config, ord_no):
    sql_exp = f'SELECT cli_no FROM order_header JOIN client USING(cli_id) WHERE ord_no = {ord_no}'
    result_set = sql_query(sql_exp, config.sigm_db_cursor)
    cli_no = scalar_data(result_set)
    return cli_no


def plq_id_cli_no(config, plq_id):
    sql_exp = f'SELECT cli_no FROM planning_lot_quantity ' \
              f'JOIN order_line USING(orl_id) ' \
              f'JOIN order_header USING(ord_id) ' \
              f'JOIN client USING(cli_id) ' \
              f'WHERE plq_id = {plq_id}'
    result_set = sql_query(sql_exp, config.sigm_db_cursor)
    cli_no = scalar_data(result_set)
    return cli_no
