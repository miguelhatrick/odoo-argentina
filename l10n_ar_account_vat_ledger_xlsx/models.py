 # -*- coding: utf-8 -*-
##############################################################################
# For copyright and license notices, see __manifest__.py file in module root
# directory
##############################################################################
from odoo import models, fields, api
from odoo.exceptions import UserError,ValidationError
from odoo.tools.safe_eval import safe_eval

from odoo.addons.report_xlsx.report.report_xlsx import ReportXlsxAbstract

import pdb
import datetime
from dateutil.parser import *
import logging
_logger = logging.getLogger(__name__)

class AccountInvoiceTax(models.Model):
    _inherit = 'account.invoice.tax'

    #line_types = fields.Char('CAE Barcode',compute=_compute_cae_barcode)

class AccountInvoice(models.Model):
    _inherit = 'account.invoice'

    @api.multi
    def _compute_vat_ledger_ref(self):
        for inv in self:
            if inv.type in ['in_invoice','in_refund']:
                for line in inv.invoice_line_ids:
                    if line.account_id:
                        inv.vat_ledger_ref = line.account_id.name


    @api.multi
    def _compute_cae_barcode(self):
        #company.partner_id.document_number,
        #o.journal_id.journal_class_id.afip_code,
        #o.journal_id.point_of_sale,
        #int(o.afip_cae or 0),
        #int(o.afip_cae_due is not False and flatdate(o.afip_cae_due) or 0)
        for inv in self:
            inv.cae_barcode = str(inv.company_id.partner_id.document_number) + str(inv.journal_id.journal_class_id.afip_code) + \
                            str(inv.journal_id.point_of_sale) + str(inv.afip_cae or 0) + str(inv.afip_cae_due or 0).replace('-','')

    cae_barcode = fields.Char('CAE Barcode',compute=_compute_cae_barcode)
    vat_ledger_ref = fields.Char('VAT Ledger Ref',compute=_compute_vat_ledger_ref)

class XHeader(object):
    def __init__(self, name, hint="", column=0, hidden=False):
        self.name = name
        self.hint = hint
        self.column = column
        self.hidden = hidden

        if (hint==""):
            self.hint = name

class VatLedgerXlsx(models.AbstractModel):
    _name = 'report.account.vat.ledger.xlsx'
    _inherit = 'report.report_xlsx.abstract'

    def generate_xlsx_report(self, workbook, data, ledgers):
        _logger.info(ledgers)
        _logger.info(ledgers.get_tax_groups_columns())
        _logger.info(ledgers.invoice_ids)
        for obj in ledgers:
            report_name = obj.name
            # One sheet by partner
            sheet = workbook.add_worksheet(report_name[:31])
            bold = workbook.add_format({'bold': True})
            sheet.write(0, 0, obj.name, bold)
            headersdef = []
            headersdef.append( XHeader( 'Fecha','Fecha' ) )
            headersdef.append( XHeader( 'Tipo','Tipo' ) )
            headersdef.append( XHeader( 'Cpbte','Comprobante' ) )
            headersdef.append( XHeader( 'Nombre',(obj.type=="sale" and 'Cliente' or 'Proveedor' ) ) )
            #headersdef.append( XHeader( 'Anul' ) )
            headersdef.append( XHeader( 'IVA','Cond. IVA' ) )
            headersdef.append( XHeader( 'CUIT','CUIT o DOC' ) )

            #headersdef.append( XHeader( 'PTIPO','Producto Tipo' ) )
            #headersdef.append( XHeader( 'Cat','Categoria' ) )
            headersdef.append( XHeader( 'Gravado' ) )
            headersdef.append( XHeader( 'No Gravado' ) )
            for column_name, taxgroup in obj.get_tax_groups_columns():
                _logger.info("column_name:"+str(column_name))
                headersdef.append( XHeader(column_name))
            headersdef.append( XHeader( 'Otros','Otros Imp.' ) )
            #headersdef.append( XHeader( 'I.V.A.' ) )
            #headersdef.append( XHeader( 'IIBB' ) )
            #headersdef.append( XHeader( 'P.Gcias.' ) )
            headersdef.append( XHeader( 'Total' ) )
            headers = {}
            index = 0
            for header in headersdef:
                if (header.hidden==False):
                    sheet.write(3,index,header.hint,bold)
                    header.column = index
                    headers[header.name] = header
                    index = index + 1
            row_index = 4+2

            for inv in obj.invoice_ids:
                inv_sign = (inv.type in ("out_refund","in_refund") and -1.0 or 1.0)
                taxes_product_types = {}
                for line in inv.invoice_line_ids:
                    txids = line.invoice_line_tax_ids
                    #_logger.info(txids)
                    pid_type = line.product_id.type
                    taxid = txids
                    if ( len(txids)>1 ):
                        taxid = txids[0]
                    if (taxid):
                        if (taxid.description in taxes_product_types):
                            prevtype = taxes_product_types[taxid.description]
                            if (prevtype!=pid_type):
                                taxes_product_types[taxid.description] = "MIXTO"
                        else:
                            taxes_product_types[taxid.description] = pid_type
                    taxes_product_types[taxid.description] = pid_type
                _logger.info(taxes_product_types)

                #_logger.info(inv.date_invoice)
                sheet.write(row_index,headers["Fecha"].column, parse(str(inv.date_invoice)).strftime("%d/%m/%Y"))
                sheet.write(row_index,headers["Tipo"].column,inv.vat_ledger_ref)
                sheet.write(row_index,headers["Cpbte"].column,inv.display_name)
                sheet.write(row_index,headers["Nombre"].column,inv.partner_id.name)
                sheet.write(row_index,headers["CUIT"].column,inv.partner_id.main_id_number)
                sheet.write(row_index,headers["IVA"].column,inv.partner_id.afip_responsability_type_id.name)
                #gravado
                vat_base_amount = inv_sign*inv.cc_vat_taxable_amount
                sheet.write(row_index,headers["Gravado"].column, vat_base_amount )
                #no gravado
                not_vat_base_amount = inv_sign*(inv.cc_amount_untaxed - inv.cc_vat_taxable_amount)
                sheet.write(row_index,headers["No Gravado"].column, round(not_vat_base_amount,2) )

                for column_name, tax_groups in obj.get_tax_groups_columns():
                    _logger.info(column_name)
                    cc_amount = inv_sign*inv.tax_line_ids.filtered(lambda x: x.tax_id.tax_group_id == tax_groups).cc_amount
                    sheet.write( row_index, headers[column_name].column, cc_amount )

                other_taxes_amount = inv_sign*(inv.cc_other_taxes_amount)
                sheet.write( row_index, headers["Otros"].column, other_taxes_amount )

                total = inv_sign*inv.cc_amount_total
                sheet.write(row_index,headers["Total"].column,total)

                row_index = row_index + 1

            # Details sheets
            sheet_details = workbook.add_worksheet("Details")
            bold = workbook.add_format({'bold': True})
            sheet_details.write(0, 0, obj.name, bold)
            headersdef = []
            headersdef.append( XHeader( 'Fecha','Fecha' ) )
            headersdef.append( XHeader( 'Tipo','Tipo' ) )
            headersdef.append( XHeader( 'Cpbte','Comprobante' ) )
            headersdef.append( XHeader( 'Nombre',(obj.type=="sale" and 'Cliente' or 'Proveedor' ) ) )
            #headersdef.append( XHeader( 'Anul' ) )
            headersdef.append( XHeader( 'IVA','Cond. IVA' ) )
            headersdef.append( XHeader( 'CUIT','CUIT o DOC' ) )

            #headersdef.append( XHeader( 'PTIPO','Producto Tipo' ) )
            #headersdef.append( XHeader( 'Cat','Categoria' ) )
            headersdef.append( XHeader( 'Gravado' ) )
            headersdef.append( XHeader( 'No Gravado' ) )
            for column_name, taxgroup in obj.get_tax_groups_columns():
                _logger.info("column_name:"+str(column_name))
                headersdef.append( XHeader(column_name))

            other_tax_columns = obj.get_other_tax_columns()

            for column_name, tax in other_tax_columns:
                headersdef.append( XHeader(column_name))

            headersdef.append( XHeader( 'Total' ) )
            headers = {}
            index = 0
            for header in headersdef:
                if (header.hidden==False):
                    sheet_details.write(3,index,header.hint,bold)
                    header.column = index
                    headers[header.name] = header
                    index = index + 1
            row_index = 4+2

            for inv in obj.invoice_ids:
                inv_sign = (inv.type in ("out_refund","in_refund") and -1.0 or 1.0)
                taxes_product_types = {}
                for line in inv.invoice_line_ids:
                    txids = line.invoice_line_tax_ids
                    #_logger.info(txids)
                    pid_type = line.product_id.type
                    taxid = txids
                    if ( len(txids)>1 ):
                        taxid = txids[0]
                    if (taxid):
                        if (taxid.description in taxes_product_types):
                            prevtype = taxes_product_types[taxid.description]
                            if (prevtype!=pid_type):
                                taxes_product_types[taxid.description] = "MIXTO"
                        else:
                            taxes_product_types[taxid.description] = pid_type
                    taxes_product_types[taxid.description] = pid_type
                _logger.info(taxes_product_types)

                #_logger.info(inv.date_invoice)
                sheet_details.write(row_index,headers["Fecha"].column, parse(str(inv.date_invoice)).strftime("%d/%m/%Y"))
                sheet_details.write(row_index,headers["Tipo"].column,inv.vat_ledger_ref)
                sheet_details.write(row_index,headers["Cpbte"].column,inv.display_name)
                sheet_details.write(row_index,headers["Nombre"].column,inv.partner_id.name)
                sheet_details.write(row_index,headers["CUIT"].column,inv.partner_id.main_id_number)
                sheet_details.write(row_index,headers["IVA"].column,inv.partner_id.afip_responsability_type_id.name)
                #gravado
                vat_base_amount = inv_sign*inv.cc_vat_taxable_amount
                sheet_details.write(row_index,headers["Gravado"].column, vat_base_amount )
                #no gravado
                not_vat_base_amount = inv_sign*(inv.cc_amount_untaxed - inv.cc_vat_taxable_amount)
                sheet_details.write(row_index,headers["No Gravado"].column, round(not_vat_base_amount,2) )

                for column_name, tax_groups in obj.get_tax_groups_columns():
                    _logger.info(column_name)
                    cc_amount = inv_sign*inv.tax_line_ids.filtered(lambda x: x.tax_id.tax_group_id == tax_groups).cc_amount
                    sheet_details.write( row_index, headers[column_name].column, cc_amount )

                # others taxes
                for column_name, tax in other_tax_columns:
                    cc_amount = inv_sign*inv.tax_line_ids.filtered(lambda x: x.tax_id.id == tax).cc_amount
                    sheet_details.write( row_index, headers[column_name].column, cc_amount )

                total = inv_sign*inv.cc_amount_total
                sheet_details.write(row_index,headers["Total"].column,total)

                row_index = row_index + 1



#VatLedgerXlsx('report.account.vat.ledger.xlsx','account.vat.ledger')
