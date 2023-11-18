#!/usr/bin/env python3

#************************************************************#
#                                                            #
#  Written by Yuri H. Galvao <yuri@galvao.ca>, January 2023  #
#                                                            #
#************************************************************#

import sys
from flask import Flask, request, url_for, render_template, jsonify
from label_generator import get_order_series

# Declaring variables and instantiating objects
args = sys.argv[1:] # List of arguments that were passed, if any

on_premises = True if '--on-premises' in args else False
yes_for_all = True if '--yes-for-all' in args else False

app = Flask(__name__)

# Defining functions - and decorating them
@app.route('/_show_invoice_info')
def show_invoice_info(error:bool=False):
    """
    Retrieves invoice information and displays it to the user.

    Arguments:
        error (bool, optional): if True, displays an error message instead of the invoice information.

    Returns:
        json: a JSON object containing the destination address, products, and attention details.
    """

    from label_generator import get_address, get_products_names # Lazy load, to prevent cold starts

    global order_n
    order_n = request.args.get('order_n', 0, type=int)
    global order_series
    order_series = get_order_series(order_n)

    if order_series is not None:
        to_address = get_address(order_series, 'to', 'no', '', '')
        to_address = to_address[0] + '\n' + to_address[1] + '\n' + to_address[2]
        products = get_products_names(order_series)
        d_name = order_series.ShipMethodRef['name'] if order_series.ShipMethodRef else ''
        delivery_name = d_name if 'pickup' not in d_name.strip().lower() and 'pick up' not in d_name.strip().lower() else ''
    else:
        error = True

    if error:
        to_address = 'ORDER NOT FOUND! CHECK ORDER STATUS ON THE BACKEND SYSTEM OR TRY AGAIN!'
        products = []
        delivery_name = ''

    return jsonify(destination_address=to_address, products=products, attn=delivery_name)

@app.route('/', methods=['GET', 'POST'])
def index():
    """
    Handles the main route of the Flask application, processes form data, and calls the label generation function.
    After that, renders the HTML template for the main page, by using the Jinja2 engine, along with the result and 
    a link to the generated PDF.
    """

    result = False
    link = ''
    global order_n
    global order_series

    if request.method == 'POST':
        form = request.form
        try:
            template = form['input_product_qty_check']
        except:
            template = ''

        try:
            products_sizes_names = form['product'].strip('"').replace(',', ', ')
            products_sizes_names = [products_sizes_names] if len(products_sizes_names) < 98 else [products_sizes_names[:97]]
        except:
            products_sizes_names = form['products'].strip('"').strip('||').split('||,')
        
        try:
            from label_generator import output_label, logging # Lazy load, to prevent cold starts
            order_n_ = int(form['order_n2'])
            if order_n != order_n_:
                raise ValueError

            result, link = output_label(
                template,
                order_series,
                order_n,
                products_sizes_names,
                form['add_info'],
                form['package_type'],
                int(form['packages_qty']),
                [int(n) for n in form['products_qty'].strip('"').split(',')],
                to_address=form['to_address'],
                additional_info_to='Attn.: ' + form['attn'] if form['attn'] not in ('', ' ', None) else ''
            )
        except Exception as e:
            logging.error(f'''Error with the order number! Exception: {repr(e)}''')
            logging.error('Trying again...')

            try:
                order_series = get_order_series(order_n_)
                result, link = output_label(
                    template,
                    order_series,
                    order_n_,
                    products_sizes_names,
                    form['add_info'],
                    form['package_type'],
                    int(form['packages_qty']),
                    [int(n) for n in form['products_qty'].strip('"').split(',')],
                    to_address=form['to_address'],
                    additional_info_to='Attn.: ' + form['attn'] if form['attn'] not in ('', ' ', None) else ''
                )
            except Exception as e:
                logging.critical(f'''Error! Is there an already-fetched order? Exception: {e}''')
                show_invoice_info(error=True)

    return render_template('index.html', result=result, link_to_pdf=link)

if __name__ == '__main__':
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
