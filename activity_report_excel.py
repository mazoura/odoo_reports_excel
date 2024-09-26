from odoo import models
from datetime import datetime
from collections import defaultdict
from datetime import timedelta
import pymssql



class CommercialReportExcel(models.AbstractModel):
    _name = 'report.mb_reporting.commercial_excel'
    _inherit = 'report.report_xlsx.abstract'
    
    def get_type_client(self,client):
        if 'DGSN' in str(client) or 'DGF' in str(client) or 'DGPC' in str(client):
            return 'Para-Militaire'
        if 'MDN' in str(client) or 'DCM' in str(client):
            return 'Militaire'
        if 'AMS' in str(client):
            return 'AMS-Tiaret'
        return 'Grand Public'  

    def generate_xlsx_report(self, workbook, data, lines):

        date_from = data['form']['date_from']
        date_to = data['form']['date_to']

        date_from = datetime.strptime(date_from, '%Y-%m-%d')
        date_to = datetime.strptime(date_to, '%Y-%m-%d')

        header = workbook.add_format({'bold': True,'font_size':20,'fg_color': 'yellow'})
        information = workbook.add_format({'bold': True,'font_size':15})
        table_header = workbook.add_format({'bold': True,'font_size':12,'border':1 ,'font_color': 'white','fg_color': 'black'})
        
        table_row = workbook.add_format({'font_size':12})
        table_styled_row = workbook.add_format({'font_size':12,'fg_color': '#D9D9D9'})

        sheet =workbook.add_worksheet("Commandes Clients")
        sheet2 =workbook.add_worksheet("Etats Des Commandes")
        sheet3 =workbook.add_worksheet("Etats Des Livraisons")
        sheet4 =workbook.add_worksheet("Chiffre d'affaire")
        
        
        sheet.set_column('A:A', 20)
        sheet.set_column('B:B', 20)
        sheet.set_column('C:C', 50)
        sheet.set_column('D:D', 30)
        sheet.set_column('E:E', 20)
        sheet.set_column('F:F', 20)
        sheet.set_column('G:G', 20)
        sheet.set_column('H:H', 30)
        sheet.set_column('I:I', 30)
        sheet.set_column('J:J', 30)
        sheet.set_column('K:K', 40)
        sheet.set_column('L:L', 40)
        sheet.set_column('M:M', 40)
        sheet.set_column('N:N', 40)
        sheet.set_column('O:O', 30)
        sheet.set_column('P:P', 30)
        sheet.set_column('Q:Q', 30)
        sheet.set_column('R:R', 30)
        sheet.set_column('S:S', 30)
        sheet.set_column('T:T', 30)
        sheet.set_column('U:U', 30)
        sheet.set_column('V:V', 30)
        sheet.set_column('W:W', 30)
        sheet.set_column('X:X', 30)
        sheet.set_column('Y:Y', 30)
        sheet.set_column('Z:Z', 30)
        sheet.set_column('AA:AA', 30)
        sheet.set_column('AB:AB', 30)   
        
        operation_type = defaultdict()
        
        operation_type["SAFAV-MB-TEST"] = 97
        operation_type["SAFAV-MB-PROD"] = 97
        operation_type["AMS-PROD"] = 145
        operation_type["SAFAV-MB-PROD-05-06"] = 97

        i = 1
        sheet.write(i, 2, "Commandes Clients du "+str(date_from.strftime("%d/%m/%Y")) + " au " + str(date_to.strftime("%d/%m/%Y")), table_header)
        
        i = 3
        #table header
        sheet.write(i, 0, "Model", table_header)
        sheet.write(i, 1, "Type", table_header)
        sheet.write(i, 2, "Client", table_header)
        sheet.write(i, 3, "Type client", table_header)
        sheet.write(i, 4, "Contrat", table_header)
        sheet.write(i, 5, "Délai de livraison", table_header)
        sheet.write(i, 6, "Date de notification", table_header)
        sheet.write(i, 7, "Date butoire", table_header)
        sheet.write(i, 8, "Montant du contrat", table_header)
        sheet.write(i, 9, "Montant du contrat", table_header)
        
        sheet.write(i, 10, "Désignation ", table_header)
        sheet.write(i, 11, "PU", table_header)
        sheet.write(i, 12, "Qte", table_header)
        sheet.write(i, 13, "Décompte", table_header)
        sheet.write(i, 14, "Qte livrée antérieurement ", table_header)
        sheet.write(i, 15, "Reste à livrer", table_header)
        sheet.write(i, 16, "Montant réalisée", table_header)
        sheet.write(i, 17, "Montant restant", table_header)
        sheet.write(i, 18, "Taux d'execution", table_header)
        sheet.write(i, 19, "Observation", table_header)
        sheet.write(i, 20, "Date actualisation", table_header)
        sheet.write(i, 21, "Statut contrat", table_header)            

        vehicle_type = defaultdict()

        vehicle_type['S13'] = 'A1'
        vehicle_type['S02'] = 'A1'
        vehicle_type['S12'] = 'A1'
        vehicle_type['S21'] = 'A1'
        #******************************26/09/24
        vehicle_type['S28'] = 'A1'

        vehicle_type['S03'] = 'A2'
        vehicle_type['S04'] = 'A2'
        vehicle_type['S14'] = 'A2'
        vehicle_type['S06'] = 'A2'
        vehicle_type['S10'] = 'A2'
        vehicle_type['S30'] = 'A2'        
        vehicle_type['S15'] = 'A2'
        vehicle_type['S19'] = 'A2'
        vehicle_type['S20'] = 'A2'
        vehicle_type['S22'] = 'A2'
        
        vehicle_type['S23'] = 'A4'
        vehicle_type['S27'] = 'A4'
        vehicle_type['S25'] = 'A4'
        vehicle_type['S09'] = 'A4'
        vehicle_type['S11'] = 'A4'
        vehicle_type['SS1'] = 'A4'       
        
        vehicle_type['G06'] = 'BA6'
        vehicle_type['G09'] = 'BA9'
       
        # ******************* sheet 2 *************************


        sheet2.set_column('A:A', 20)
        sheet2.set_column('B:B', 20)
        sheet2.set_column('C:C', 50)
        sheet2.set_column('D:D', 50)
        sheet2.set_column('E:E', 50)
        sheet2.set_column('F:F', 50)
        sheet2.set_column('G:G', 50)
        
        j = 1
        sheet2.write(j, 2, "Etats Des Commandes du "+str(date_from.strftime("%d/%m/%Y")) + " au " + str(date_to.strftime("%d/%m/%Y")), table_header)
        
        j = 3

        sheet2.write(j, 0, "Produit", table_header)
        sheet2.write(j, 1, "Qte Global", table_header)
        sheet2.write(j, 2, "Montant Global", table_header)
        sheet2.write(j, 3, "Qte Livrée Antérieurement", table_header)
        sheet2.write(j, 4, "Montant Réalisé", table_header)
        sheet2.write(j, 5, "Reste à Livrer", table_header)
        sheet2.write(j, 6, "Montant Reste", table_header)

        qte_global = defaultdict()
        montant_global = defaultdict()
        qte_livrer = defaultdict()
        montant_realise = defaultdict()
        qte_reste = defaultdict()
        montant_reste = defaultdict()

        qte_global['BA6'] = 0
        qte_global['BA9'] = 0
        qte_global['A1'] = 0
        qte_global['A2'] = 0
        qte_global['A4'] = 0

        montant_global['BA6'] = 0
        montant_global['BA9'] = 0
        montant_global['A1'] = 0
        montant_global['A2'] = 0
        montant_global['A4'] = 0
        
        qte_livrer['BA6'] = 0
        qte_livrer['BA9'] = 0
        qte_livrer['A1'] = 0
        qte_livrer['A2'] = 0
        qte_livrer['A4'] = 0

        montant_realise['BA6'] = 0
        montant_realise['BA9'] = 0
        montant_realise['A1'] = 0
        montant_realise['A2'] = 0
        montant_realise['A4'] = 0

        qte_reste['BA6'] = 0
        qte_reste['BA9'] = 0
        qte_reste['A1'] = 0
        qte_reste['A2'] = 0
        qte_reste['A4'] = 0

        montant_reste['BA6'] = 0
        montant_reste['BA9'] = 0
        montant_reste['A1'] = 0
        montant_reste['A2'] = 0
        montant_reste['A4'] = 0
   

        # ******************* sheet 2 *************************

        i = 4
        # ******************* contrats ********************************
      

        docs =self.env['account.analytic.account'].search([('date_start', '>=', str(date_from)),('date_end', '<', str(date_to))])
        for doc in docs:
            sheet.write(i, 8,  "{:,.2f} ".format(doc.amount), table_row)
            
            for line in doc.recurring_invoice_line_ids:
                
                model = ''
                code = line.product_id.product_code
                if code[2:5]== 'G06':
                    sheet.write(i, 0, 'G-CLASS', table_row)
                    sheet.write(i, 1, 'BA6', table_row)
                    model = 'BA6'
                if code[2:5]== 'G09':
                    sheet.write(i, 0, 'G-CLASS', table_row)
                    sheet.write(i, 1, 'BA9', table_row)
                    model = 'BA9'

                if 'S' in code[2:5]:
                    sheet.write(i, 0, 'SPRINTER', table_row)
                    sheet.write(i, 1, vehicle_type[code[2:5]], table_row)
                    model = vehicle_type[code[2:5]]

                client = doc.partner_id.name
                sheet.write(i, 2, client, table_row)
                type_client = 'Grand Public'
                if 'DGSN' in str(client) or 'DGF' in str(client) or 'DGPC' in str(client):
                    type_client = 'Para-Militaire'
                if 'MDN' in str(client) or 'DCM' in str(client):
                    type_client = 'Militaire'
                if 'AMS' in str(client):
                    type_client = 'AMS-Tiaret'              
                sheet.write(i, 3, type_client, table_row)

                
                sheet.write(i, 4, doc.code, table_row)                
                sheet.write(i, 5, doc.delivery_date.strftime("%d/%m/%Y"), table_row)
                
                if doc.date_start:
                    sheet.write(i, 6, doc.date_start.strftime("%d/%m/%Y"), table_row)
                if doc.date_end:
                    sheet.write(i, 7, doc.date_end.strftime("%d/%m/%Y"), table_row)

                

                line_amount = line.price_unit * line.quantity
                sheet.write(i, 9, "{:,.2f}".format(line_amount) , table_row)

                sheet.write(i, 10, line.product_id.name , table_row)
                sheet.write(i, 11, "{:,.2f} ".format(line.price_unit) , table_row)
                sheet.write(i, 12, line.quantity , table_row)
                sheet.write(i, 13, "{:,.2f}".format(line_amount) , table_row)


                # ********************* livraison ****************************
                s = 0
                bls = self.env['stock.picking'].search([('contract_id', '=', doc.id),('picking_type_id', '=',  operation_type[self.env.cr.dbname]),('state','in',['done','partially_cancel'])])
                for bl in bls:
                    for move in bl.move_ids_without_package:
                        if move.product_id.id == line.product_id.id:
                            if bl.state == 'partially_cancel':
                                for move_line in move.move_line_ids:
                                    if move_line.state not in ('cancel', 'draft'):
                                        s+= move_line.qty_done
                            else:
                                s += move.product_uom_qty
                            break
                
                sheet.write(i, 14, s , table_row)
                sheet.write(i, 15, line.quantity - s , table_row)
                sheet.write(i, 16, "{:,.2f} ".format(line.price_unit * s) , table_row)
                sheet.write(i, 17, "{:,.2f} ".format(line.price_unit * (line.quantity - s)) , table_row)
                if line.quantity !=0:
                    sheet.write(i, 18, str(s * 100 / line.quantity) + "%", table_row)
                sheet.write(i, 19, '', table_row)
                sheet.write(i, 20, '', table_row)

                if line.quantity == s:
                    sheet.write(i, 21, 'Contrat Cloturé', table_row)
                else:
                    sheet.write(i, 21, 'Contrat En-cours', table_row)

                # ********************* livraison ****************************

                
                i += 1
                if model != '':
                    qte_global[model] += line.quantity
                    montant_global[model] += line.price_unit * line.quantity
                    qte_livrer[model] += s
                    montant_realise[model] += line.price_unit * s
                    qte_reste[model] += line.quantity - s
                    montant_reste[model] += (line.quantity - s) * line.price_unit


        # ******************* commandes ********************************
        docs = self.env['sale.order'].search([('state', '=', 'sale'),('confirmation_date', '>=', str(date_from)),('confirmation_date', '<', str(date_to))])
        for doc in docs:
            sheet.write(i, 8,  "{:,.2f} ".format(doc.amount_untaxed), table_row)
            
            for line in doc.order_line:

                model = ''
                code = line.product_id.product_code
                if code[2:5]== 'G06':
                    sheet.write(i, 0, 'G-CLASS', table_row)
                    sheet.write(i, 1, 'BA6', table_row)
                    model = 'BA6'
                if code[2:5]== 'G09':
                    sheet.write(i, 0, 'G-CLASS', table_row)
                    sheet.write(i, 1, 'BA9', table_row)
                    model = 'BA9'

                if 'S' in code[2:5]:
                    sheet.write(i, 0, 'SPRINTER', table_row)
                    sheet.write(i, 1, vehicle_type[code[2:5]], table_row)
                    model = vehicle_type[code[2:5]]
                    print('*********************** '+str(model))

                client = doc.partner_id.name
                sheet.write(i, 2, client, table_row)
                type_client = 'Grand Public'
                if 'DGSN' in str(client) or 'DGF' in str(client) or 'DGPC' in str(client):
                    type_client = 'Para-Militaire'
                if 'MDN' in str(client) or 'DCM' in str(client):
                    type_client = 'Militaire'
                if 'AMS' in str(client):
                    type_client = 'AMS-Tiaret'              
                sheet.write(i, 3, type_client, table_row)

                if doc.is_external:
                    sheet.write(i, 4, doc.external_ref, table_row)
                else:
                    sheet.write(i, 4, doc.name, table_row)

                sheet.write(i, 5, doc.delay, table_row)
                sheet.write(i, 6, doc.confirmation_date.strftime("%d/%m/%Y"), table_row)
                date_but = doc.confirmation_date + timedelta(days= int(doc.delay))
                sheet.write(i, 7, date_but.strftime("%d/%m/%Y"), table_row)

                

                line_amount = line.price_unit * line.product_uom_qty
                sheet.write(i, 9, "{:,.2f}".format(line_amount) , table_row)

                sheet.write(i, 10, line.name , table_row)
                sheet.write(i, 11, "{:,.2f} ".format(line.price_unit) , table_row)
                sheet.write(i, 12, line.product_uom_qty , table_row)
                sheet.write(i, 13, "{:,.2f}".format(line_amount) , table_row)

              

                # ********************* livraison ****************************
                s = 0
                bls = self.env['stock.picking'].search([('commande', '=', doc.id),('picking_type_id', '=',  operation_type[self.env.cr.dbname]),('state','in',['done','partially_cancel'])])
                for bl in bls:
                    for move in bl.move_ids_without_package:
                        if move.product_id.id == line.product_id.id:
                            if bl.state == 'partially_cancel':
                                for move_line in move.move_line_ids:
                                    if move_line.state not in ('cancel', 'draft'):
                                        s+= 1
                            else:
                                s += move.product_uom_qty
                            break
                
                sheet.write(i, 14, s , table_row)
                sheet.write(i, 15, line.product_uom_qty - s , table_row)
                sheet.write(i, 16, "{:,.2f} ".format(line.price_unit * s) , table_row)
                sheet.write(i, 17, "{:,.2f} ".format(line.price_unit * (line.product_uom_qty - s)) , table_row)
                if line.product_uom_qty !=0:
                    sheet.write(i, 18, str(s * 100 / line.product_uom_qty) + "%", table_row)
                sheet.write(i, 19, '', table_row)
                sheet.write(i, 20, '', table_row)

                if line.product_uom_qty == s:
                    sheet.write(i, 21, 'Contrat Cloturé', table_row)
                else:
                    sheet.write(i, 21, 'Contrat En-cours', table_row)

                # ********************* livraison ****************************
                if model != '':
                    qte_global[model] += line.product_uom_qty
                    montant_global[model] += line.price_unit * line.product_uom_qty
                    qte_livrer[model] += s
                    montant_realise[model] += line.price_unit * s
                    qte_reste[model] += line.product_uom_qty - s
                    montant_reste[model] += (line.product_uom_qty - s) * line.price_unit

                i += 1

        j = 4

        classe_g = defaultdict()
        sprinter = defaultdict()
        total = defaultdict()

        classe_g['qte_global'] = qte_global['BA6'] + qte_global['BA9']
        classe_g['montant_global'] = montant_global['BA6'] + montant_global['BA9']
        classe_g['qte_livrer'] = qte_livrer['BA6'] + qte_livrer['BA9']
        classe_g['montant_realise'] = montant_realise['BA6'] + montant_realise['BA9'] 
        classe_g['qte_reste'] = qte_reste['BA6'] + qte_reste['BA9']
        classe_g['montant_reste'] = montant_reste['BA6'] + montant_reste['BA9']


        sprinter['qte_global'] = qte_global['A1'] + qte_global['A2'] + qte_global['A4']
        sprinter['montant_global'] = montant_global['A1'] + montant_global['A2'] + qte_global['A4']
        sprinter['qte_livrer'] = qte_livrer['A1'] + qte_livrer['A2'] + qte_livrer['A4']
        sprinter['montant_realise'] = montant_realise['A1'] + montant_realise['A2'] + qte_global['A4']
        sprinter['qte_reste'] = qte_reste['A1'] + qte_reste['A2'] + qte_reste['A4']
        sprinter['montant_reste'] = montant_reste['A1'] + montant_reste['A2'] + qte_global['A4']

        total['qte_global'] = classe_g['qte_global'] + sprinter['qte_global']
        total['montant_global'] = classe_g['montant_global'] + sprinter['montant_global']
        total['qte_livrer'] = classe_g['qte_livrer'] + sprinter['qte_livrer']
        total['montant_realise'] = classe_g['montant_realise'] + sprinter['montant_realise']
        total['qte_reste'] = classe_g['qte_reste'] + sprinter['qte_reste']
        total['montant_reste'] = classe_g['montant_reste'] + sprinter['montant_reste']

        for model in ["CLASSE-G","BA6","BA9","SPRINTER","A1","A2","A4","TOTAL"]:
            
            if model == 'CLASSE-G':
                sheet2.write(j, 0, model, table_styled_row)
                sheet2.write(j, 1, classe_g['qte_global'], table_styled_row)
                sheet2.write(j, 2, "{:,.2f} ".format(classe_g['montant_global']), table_styled_row)
                sheet2.write(j, 3, classe_g['qte_livrer'], table_styled_row)
                sheet2.write(j, 4, "{:,.2f} ".format(classe_g['montant_realise']), table_styled_row)
                sheet2.write(j, 5, classe_g['qte_reste'], table_styled_row)
                sheet2.write(j, 6, "{:,.2f} ".format(classe_g['montant_reste']), table_styled_row)
            
            if model == 'SPRINTER':
                sheet2.write(j, 0, model, table_styled_row)
                sheet2.write(j, 1, sprinter['qte_global'], table_styled_row)
                sheet2.write(j, 2, "{:,.2f} ".format(sprinter['montant_global']), table_styled_row)
                sheet2.write(j, 3, sprinter['qte_livrer'], table_styled_row)
                sheet2.write(j, 4, "{:,.2f} ".format(sprinter['montant_realise']), table_styled_row)
                sheet2.write(j, 5, sprinter['qte_reste'], table_styled_row)
                sheet2.write(j, 6, "{:,.2f} ".format(sprinter['montant_reste']), table_styled_row)

            if model == 'TOTAL':
                sheet2.write(j, 0, model, table_header)
                sheet2.write(j, 1, total['qte_global'], table_header)
                sheet2.write(j, 2, "{:,.2f} ".format(total['montant_global']), table_header)
                sheet2.write(j, 3, total['qte_livrer'], table_header)
                sheet2.write(j, 4, "{:,.2f} ".format(total['montant_realise']), table_header)
                sheet2.write(j, 5, total['qte_reste'], table_header)
                sheet2.write(j, 6, "{:,.2f} ".format(total['montant_reste']), table_header)
            
            if model in ["BA6","BA9","A1","A2","A4"]:
                sheet2.write(j, 0, model, table_row)
                sheet2.write(j, 1, qte_global[model], table_row)
                sheet2.write(j, 2, "{:,.2f} ".format(montant_global[model]), table_row)
                sheet2.write(j, 3, qte_livrer[model], table_row)
                sheet2.write(j, 4, "{:,.2f} ".format(montant_realise[model]), table_row)
                sheet2.write(j, 5, qte_reste[model], table_row)
                sheet2.write(j, 6, "{:,.2f} ".format(montant_reste[model]), table_row)
            
            j += 1      
        
        
        
        # ********************************** ETAT DES LIVRAISON ************************************                           

        sheet3.set_column('A:A', 20)
        sheet3.set_column('B:B', 20)
        sheet3.set_column('C:C', 80)
        sheet3.set_column('D:D', 10)
        sheet3.set_column('E:E', 10)

        sheet3.set_column('F:F', 10)
        sheet3.set_column('G:G', 10)
        sheet3.set_column('H:H', 10)
        sheet3.set_column('I:I', 10)
        sheet3.set_column('J:J', 10)
        sheet3.set_column('K:K', 10)
        sheet3.set_column('L:L', 10)
        sheet3.set_column('M:M', 10)
        sheet3.set_column('N:N', 10)
        sheet3.set_column('O:O', 10)        
        sheet3.set_column('P:P', 20)
        
        i = 1
        sheet3.write(i, 2, "Etat des livraisons du "+str(date_from.strftime("%d/%m/%Y")) + " au " + str(date_to.strftime("%d/%m/%Y")), table_header)
        i = 3
        #table header
        sheet3.write(i, 0, "Model", table_header)
        sheet3.write(i, 1, "Type", table_header)
        sheet3.write(i, 2, "Appelation produit de Commercial", table_header)

        sheet3.write(i, 3, "Janv", table_header)
        sheet3.write(i, 4, "Févr", table_header)
        sheet3.write(i, 5, "Mars", table_header)
        sheet3.write(i, 6, "Avr", table_header)
        sheet3.write(i, 7, "Mai", table_header)
        sheet3.write(i, 8, "Juin", table_header)
        sheet3.write(i, 9, "Jui", table_header)
        sheet3.write(i, 10, "Août", table_header)
        sheet3.write(i, 11, "Sep", table_header)
        sheet3.write(i, 12, "Oct", table_header)
        sheet3.write(i, 13, "Nov", table_header)
        sheet3.write(i, 14, "Déc", table_header)
        
        
        sheet3.merge_range(i-1, 3, i-1, 14, '2021', table_header)
        sheet3.merge_range(i-1, 15, i, 15, 'Total général', table_header)

        date_index = defaultdict()
        
        date_index["01"] = 0
        date_index["02"] = 1
        date_index["03"] = 2
        date_index["04"] = 3
        date_index["05"] = 4
        date_index["06"] = 5
        date_index["07"] = 6
        date_index["08"] = 7
        date_index["09"] = 8
        date_index["10"] = 9
        date_index["11"] = 10
        date_index["12"] = 11

        vehicles = defaultdict(list)
        docs =self.env['stock.move'].search([('picking_id.picking_type_id','=', operation_type[self.env.cr.dbname]),('picking_id.scheduled_date', '>=', str(date_from)),
            ('picking_id.scheduled_date', '<', str(date_to)),('picking_id.state','in',['done','partially_cancel'])])
        
        for doc in docs:
            if doc.product_id.product_code[0:2]=='K3':
                if doc.product_id.product_code in vehicles.keys():
                    month = doc.create_date.strftime("%m")
                    if doc.picking_id.state == 'partially_cancel':
                        s = 0
                        for move_line in doc.move_line_ids:
                            if move_line.state not in ('cancel', 'draft'):
                                s+= move_line.qty_done
                        vehicles[doc.product_id.product_code][date_index[month]] += s                  
                    else:
                        vehicles[doc.product_id.product_code][date_index[month]] += doc.product_uom_qty
                    
                else:
                    vehicles[doc.product_id.product_code] = [0,0,0,0,0,0,0,0,0,0,0,0]
                    month = doc.create_date.strftime("%m")
                    if doc.picking_id.state == 'partially_cancel':
                        s = 0
                        for move_line in doc.move_line_ids:
                            if move_line.state not in ('cancel', 'draft'):
                                s+= move_line.qty_done
                        vehicles[doc.product_id.product_code][date_index[month]] += s                  
                    else:
                        vehicles[doc.product_id.product_code][date_index[month]] += doc.product_uom_qty
                    
        
        i=4
        for key in vehicles.keys():                
            code = key
            if code[2] == "S":
                sheet3.write(i, 0, "VS30", table_row)
            else:
                sheet3.write(i, 0, 'G-CLASS', table_row)
            if code[2:5] in vehicle_type.keys():
                sheet3.write(i, 1, vehicle_type[code[2:5]], table_row)
            doc = self.env['product.template'].search([('product_code','=',code)])
            sheet3.write(i, 2, doc[0].name, table_row)
            sheet3.write(i, 3, vehicles[key][0], table_row)
            sheet3.write(i, 4, vehicles[key][1], table_row)
            sheet3.write(i, 5, vehicles[key][2], table_row)
            sheet3.write(i, 6, vehicles[key][3], table_row)
            sheet3.write(i, 7, vehicles[key][4], table_row)
            sheet3.write(i, 8, vehicles[key][5], table_row)
            sheet3.write(i, 9, vehicles[key][6], table_row)
            sheet3.write(i, 10, vehicles[key][7], table_row)
            sheet3.write(i, 11, vehicles[key][8], table_row)
            sheet3.write(i, 12, vehicles[key][9], table_row)
            sheet3.write(i, 13, vehicles[key][10], table_row)
            sheet3.write(i, 14, vehicles[key][11], table_row)
            sheet3.write(i, 15, sum(vehicles[key][:]), table_header)
            i+=1

        # ********************************** ETAT DES LIVRAISON ************************************                           

        sheet4.set_column('A:A', 20)
        sheet4.set_column('B:B', 20)
        sheet4.set_column('C:C', 80)
        sheet4.set_column('D:D', 10)
        sheet4.set_column('E:E', 10)
        sheet4.set_column('F:F', 10)
        sheet4.set_column('G:G', 10)
        sheet4.set_column('H:H', 10)
        sheet4.set_column('I:I', 10)
        sheet4.set_column('J:J', 10)
        sheet4.set_column('K:K', 10)
        sheet4.set_column('L:L', 10)
        sheet4.set_column('M:M', 10)
        sheet4.set_column('N:N', 10)
        sheet4.set_column('O:O', 10)        
        sheet4.set_column('P:P', 20)        
        
        i = 1
        sheet4.write(i, 2, "Chiffre d'affaire du "+str(date_from.strftime("%d/%m/%Y")) + " au " + str(date_to.strftime("%d/%m/%Y")), table_header)
        i = 3
        #table header
       
        sheet4.write(i, 0, "Etat de produit", table_header)
        sheet4.write(i, 1, "Quantité", table_header)
        sheet4.write(i, 2, "TOTAL HT", table_header)
        
        i += 1

        delivered_invoiced_qty = defaultdict()
        delivered_not_invoiced_qty = defaultdict()
        delivered_invoiced_total = defaultdict()
        delivered_not_invoiced_total = defaultdict()
       

        del_inv_qty = 0
        del_not_inv_qty = 0
        not_del_inv_qty = 0

        del_inv_price = 0
        del_not_inv_price = 0    
        not_del_inv_price = 0

        """  docs = self.env['stock.move'].search([('picking_id.picking_type_id','=', operation_type[self.env.cr.dbname]),('picking_id.scheduled_date', '>=', str(date_from)),('picking_id.scheduled_date', '<', str(date_to)),('picking_id.state','in',['done','partially_cancel'])])
        for doc in docs:
            sol = self.env['sale.order.line'].search([('order_id','=',doc.picking_id.commande.id),('product_id','=',doc.product_id.id)])
            if doc.picking_id.state == 'partially_cancel':
                for move_line in doc.move_line_ids:
                    if move_line.state not in ('cancel', 'draft'):
                        if doc.picking_id.ref_facture:
                            if doc.picking_id.ref_facture.date_invoice >= str(date_from) and doc.picking_id.ref_facture.date_invoice < str(date_to):
                                del_inv_qty += move_line.qty_done
                                del_inv_price += move_line.qty_done * sol.price_unit
                        else:
                            del_not_inv_qty += move_line.qty_done
                            del_not_inv_price += move_line.qty_done * sol.price_unit
            else:
                if doc.picking_id.ref_facture:
                    if doc.picking_id.ref_facture.date_invoice >= datetime.date(date_from) and doc.picking_id.ref_facture.date_invoice < datetime.date(date_to):
                        del_inv_qty += doc.product_uom_qty
                        del_inv_price +=  doc.product_uom_qty * sol.price_unit
                else:
                    del_not_inv_qty += doc.product_uom_qty
                    del_not_inv_price += doc.product_uom_qty * sol.price_unit """
        

              
        docs = self.env['account.invoice'].search([('type','=','out_invoice'),('date_invoice', '>=', str(date_from)),('date_invoice', '<', str(date_to)),('state','in',['open','paid'])])
        for doc in docs:
            if doc.picking_count == 0:
                for line in doc.invoice_line_ids:
                    not_del_inv_qty += line.quantity
                not_del_inv_price += doc.amount_total_signed

            else:
                
                price_dict={}
               
                total_qty_inv = 0
                total_qty_pick = 0

                for line in doc.invoice_line_ids:
                    total_qty_inv += line.quantity
                    price_dict[line.product_id.id] = line.price_unit

                for picking in doc.invoice_picking_id:
                    
                    if picking.state in ["done", "partially_cancel"]:
                        for move in picking.move_ids_without_package:
                            total_qty_pick += move.quantity_done
                        
                        
                if total_qty_pick !=0:
                    # if total_qty_inv == total_qty_pick:
                    del_inv_price += doc.amount_total_signed
                    del_inv_qty += total_qty_pick                            
                    """ else:
                        for picking in doc.invoice_picking_id:                    
                            if picking.state in ["done", "partially_cancel"]:
                                for move in picking.move_ids_without_package:
 """





        sheet4.write(i, 0, "Livré/Facturé", table_header)
        sheet4.write(i, 1, del_inv_qty, table_row)
        sheet4.write(i, 2, "{:,.2f}".format(del_inv_price), table_row)
        i += 1
        sheet4.write(i, 0, "Livré/Non Facturé", table_header)
        sheet4.write(i, 1, del_not_inv_qty, table_row)
        sheet4.write(i, 2, "{:,.2f}".format(del_not_inv_price), table_row)
        i += 1
        sheet4.write(i, 0, "Non Livré/Facturé", table_header)
        sheet4.write(i, 1, not_del_inv_qty, table_row)
        sheet4.write(i, 2, "{:,.2f}".format(not_del_inv_price), table_row)
        i += 1
        sheet4.write(i, 0, "Chiffre d'affaire", table_header)
        sheet4.write(i, 1, del_inv_qty + del_not_inv_qty + not_del_inv_qty, table_row)
        sheet4.write(i, 2, "{:,.2f}".format(del_inv_price + del_not_inv_price + not_del_inv_price), table_row)

    # ************************************ SPRINTER *****************************************
    # ********************************** STOCK USINE ************************************

class SStockUsineExcel(models.AbstractModel):
    _name = 'report.mb_reporting.sprinter_stock_usine_excel'
    _inherit = 'report.report_xlsx.abstract'
    
    def generate_xlsx_report(self, workbook, data, lines):

        date_from = data['form']['date_from']
        date_from = datetime.strptime(date_from, '%Y-%m-%d')       
        
        cp = defaultdict()
        vin = defaultdict(list)
        
        s = defaultdict(dict)
        st = defaultdict() 
        variant1 = defaultdict()
        variant2 = defaultdict()      

        k = defaultdict(lambda: defaultdict(lambda: defaultdict(dict)))
        structure = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))

        a1t = defaultdict(list)
        a1v = defaultdict(list)    

        a2t = defaultdict(list)
        a2v = defaultdict(list)      

        a4t = defaultdict(list)
        a4v = defaultdict(list)      

        a1t = {'Fourgon Tolé S13 Blanc arctique':['A1 Cabine Approfondie','Fourgon Tolé S13 Blanc arctique']}
        a1v = {'Fourgon Vitré S02 Blanc arctique':['A1 BUS 8+1','A1 Cabine Approfondie','A1 Cellulaire -Police','Fourgon Vitré S02 Blanc arctique'],'Fourgon Vitré S12 Blanc arctique':['A1 BUS 8+1','A2 BUS 14+1','Fourgon Vitré S12 Blanc arctique'],'Fourgon Vitré S12 Gris Noir':['A1 BUS 8+1','Fourgon Vitré S12 Gris Noir']}

        a2t = {'Fourgon Tolé S04 Blanc arctique':['A2 Atelier Mobile','A2 Transport De Détenus','Fourgon Tolé S04 Blanc arctique'],'Fourgon Tolé S14 Noir':['Fourgon Tolé S14 Noir'],'Fourgon Tolé S30 Blanc arctique':['A2 Nacelle -14M','Fourgon Tolé S30 Blanc arctique']}
        a2v = {'Fourgon Vitré S06 Blanc arctique':['A2 BUS 11+1 -VIP','A2 BUS 14+1','Fourgon Vitré S06 Blanc arctique'],'Fourgon Vitré S06 Gris Noir':['A2 BUS 14+1','Fourgon Vitré S06 Gris Noir'],'Fourgon Vitré S10 Blanc arctique':['A2 Ambulance Animalière','A2 Ambulance Détenus','A2 Ambulance Médicalisé','A2 Ambulance Sanitaire','A2 Traveaux De Recherche','A2 Transport De Détenus','Fourgon Vitré S10 Blanc arctique'],'Fourgon Vitré S22 Blanc arctique':['Fourgon Vitré S22 Blanc arctique']}
        
        a4t = {'Fourgon Tolé S23 Gris Bleu':['Fourgon Tolé S23 Gris Bleu'],'Fourgon Tolé S25 Blanc arctique':['A4 Intervention -Gendarmerie','Fourgon Tolé S25 Blanc arctique'],'Fourgon Tolé S27 Blanc arctique':['Fourgon Tolé S27 Blanc arctique']}
        a4v = {'Fourgon Vitré S09 Blanc arctique':['A4 BUS 20+1 -VIP','A4 BUS 23+1','A2 BUS 14+1','A4 BUS 29+1 -Ecolier','Fourgon Vitré S09 Blanc arctique'],'Fourgon Vitré S09 Gris Noir':['A4 BUS 20+1 -VIP','A4 BUS 23+1','Fourgon Vitré S09 Gris Noir'],'Fourgon Vitré S11 Jaune Genêt':['A4 BUS 23+1','A4 BUS 29+1 -Ecolier','A4 BUS 23+1 -Ecolier','Fourgon Vitré S11 Jaune Genêt']}



        structure['A1'] = {'A1 Tolé':a1t,'Total A1 Tolé':{},'A1 Vitré':a1v,'Total A1 Vitré':{},}
        structure['Total A1'] = {}
        structure['A2'] = {'A2 Tolé':a2t,'Total A2 Tolé':{},'A2 Vitré':a2v,'Total A2 Vitré':{},}
        structure['Total A2'] = {}
        structure['A4'] = {'A4 Tolé':a4t,'Total A4 Tolé':{},'A4 Vitré':a4v,'Total A4 Vitré':{},}
        structure['Total A4'] = {}
        structure['Total Général'] = {}

        variant1['S13']='Total A1 Tolé'
        variant1['S02']='Total A1 Vitré'
        variant1['S12']='Total A1 Vitré'

        variant1['S04']='Total A2 Tolé'
        variant1['S14']='Total A2 Tolé'
        variant1['S30']='Total A2 Tolé'
        variant1['S06']='Total A2 Vitré'
        variant1['S10']='Total A2 Vitré'
        variant1['S22']='Total A2 Vitré'
        
        variant1['S23']='Total A4 Tolé'
        variant1['S25']='Total A4 Tolé'
        variant1['S27']='Total A4 Tolé'
        variant1['S09']='Total A4 Vitré'
        variant1['S11']='Total A4 Vitré'       

        variant2['S13']='Total A1'
        variant2['S02']='Total A1'
        variant2['S12']='Total A1'

        variant2['S04']='Total A2'
        variant2['S14']='Total A2'
        variant2['S30']='Total A2'
        variant2['S06']='Total A2'
        variant2['S10']='Total A2'
        variant2['S22']='Total A2'
        
        variant2['S23']='Total A4'
        variant2['S25']='Total A4'
        variant2['S27']='Total A4'
        variant2['S09']='Total A4'
        variant2['S11']='Total A4'     

        host = '192.168.1.34'
        database = 'TPR_Reports'
        user = 'safavmb\\odoo'
        password = 'Odoo@18'

        
        try:
            DB = pymssql.connect(host=host,user=user,password=password,database=database)
            cursor = DB.cursor()
        except KeyError:
                pass


        cp[5610] = 'A1 BUS 8+1'
        cp[5611] = 'A1 BUS 8+1'
        cp[5612] = 'A2 BUS 11+1 -VIP'
        cp[5613] = 'A2 BUS 14+1'
        cp[5614] = 'A2 BUS 14+1 4x4'
        cp[5615] = 'A4 BUS 20+1 -VIP'
        cp[5616] = 'A4 BUS 23+1'
        cp[5617] = 'A4 BUS 23+1 -Ecolier'
        cp[5618] = 'A4 BUS 23+1 4x4'
        cp[5619] = 'A4 BUS 29+1 -Ecolier'
        cp[5630] = 'A1 Cellulaire -Gendarmerie'
        cp[5631] = 'A1 Cellulaire -Police'
        cp[5632] = 'A1 VIP  Luxe 4+1'
        cp[5634] = 'A2 VIP  Luxe 6+2+1'
        cp[5640] = 'A4 Intervention -Gendarmerie'
        cp[5641] = 'A4 Intervention -Police'        
        cp[5650] = 'A2 Ambulance Sanitaire'
        cp[5651] = 'A2 Ambulance Sanitaire 4x4'
        cp[5652] = 'A2 Ambulance Médicalisé'
        cp[5653] = 'A2 Ambulance Médicalisé 4x4'
        cp[5654] = 'A2 Ambulance Détenus'
        cp[5655] = 'A2 Ambulance Animalière'     
        cp[5680] = 'A1 Atelier Mobile'
        cp[5681] = 'A2 Atelier Mobile'        
        cp[5690] = 'A2 Cabine Approfondie'
        cp[5691] = 'A1 Cabine Approfondie'        
        cp[5700] = 'A2 Nacelle -12M'        
        cp[5701] = 'A2 Nacelle -14M'
        cp[5710] = 'A2 Frigo'
        cp[5711] = 'A2 Mortuaire'
        cp[5712] = 'A2 Transport De Chiens'
        cp[5713] = 'A2 Transport De Fonds'
        cp[5714] = 'A2 Intervention - Police'
        cp[5715] = 'A2 Transport De Détenus'        
        
        index = defaultdict()

        index['k0'] = 4
        index['k1'] = 5
        index['k2'] = 6
        index['kc'] = 7
        index['k3'] = 8
        index['Total Général'] = 9
      
        cursor.callproc("[Reports].[spStockUsineSprinter]",[date_from + timedelta(days=1),])
        rows = list(cursor)
        for row in rows:
            
            if len(row[2]) == 12:
                if str(row[4]) not in ['UNKNOWN',None]:
                    color = str(row[4])           

            if row[11] == 'k1-kc':
                if row[9]:
                    area = 'kc'
                else:
                    area = 'k1'
            else:
                area = row[11]

            vin[row[1]] = ["","",""]
            vin[row[1]][0] = area
            try:
                vin[row[1]][1] = 'Fourgon ' + variant1[row[2][len(row[2])-3:]][5:] + ' (' + row[2][len(row[2])-3:] + ') ' + color
            except KeyError:
                pass
            #vin[row[1]][1] = 'Fvariant'

            if row[11] != 'delivered to customer':
                
                

                try:
                    s[variant1[row[2][len(row[2])-3:]]][area] += 1
                except KeyError:
                    try:
                        s[variant1[row[2][len(row[2])-3:]]][area] = 0
                        s[variant1[row[2][len(row[2])-3:]]][area] += 1
                    except KeyError:
                        pass
                    
                
                try:
                    s[variant1[row[2][len(row[2])-3:]]]['Total Général'] += 1
                except KeyError:
                    try:
                        s[variant1[row[2][len(row[2])-3:]]]['Total Général'] = 0
                        s[variant1[row[2][len(row[2])-3:]]]['Total Général'] += 1
                    except KeyError:
                        pass
                
                try:
                    s[variant2[row[2][len(row[2])-3:]]][area] += 1
                except KeyError:
                    try:
                        s[variant2[row[2][len(row[2])-3:]]][area] = 0
                        s[variant2[row[2][len(row[2])-3:]]][area] += 1
                    except KeyError:
                        pass
                
                try:
                    s[variant2[row[2][len(row[2])-3:]]]['Total Général'] += 1
                except KeyError:
                    try:
                        s[variant2[row[2][len(row[2])-3:]]]['Total Général'] = 0
                        s[variant2[row[2][len(row[2])-3:]]]['Total Général'] += 1
                    except KeyError:
                        pass
                
                try:
                    st[area] += 1
                except KeyError:
                    st[area] = 0
                    st[area] += 1
                
                try:
                    st['Total Général'] += 1
                except KeyError:
                    st['Total Général'] = 0
                    st['Total Général'] += 1
            
            

                if row[9] != None:
                    try:
                        k[area][row[2][len(row[2])-3:]][color][cp[row[9]]] += 1
                    except KeyError:
                        k[area][row[2][len(row[2])-3:]][color][cp[row[9]]] = 0
                        k[area][row[2][len(row[2])-3:]][color][cp[row[9]]] += 1
                    
                    
                    try:
                        k['s'][row[2][len(row[2])-3:]][color][cp[row[9]]] += 1
                    except KeyError:
                        k['s'][row[2][len(row[2])-3:]][color][cp[row[9]]] = 0
                        k['s'][row[2][len(row[2])-3:]][color][cp[row[9]]] += 1
                    
                    vin[row[1]][2] = cp[row[9]]

                else:
                    try:
                        k[area][row[2][len(row[2])-3:]][color]['base'] += 1
                    except KeyError:
                        k[area][row[2][len(row[2])-3:]][color]['base'] = 0
                        k[area][row[2][len(row[2])-3:]][color]['base'] += 1
                    
                    
                    try:
                        k['s'][row[2][len(row[2])-3:]][color]['base'] += 1
                    except KeyError:
                        k['s'][row[2][len(row[2])-3:]][color]['base'] = 0
                        k['s'][row[2][len(row[2])-3:]][color]['base'] += 1
  
       
        header = workbook.add_format({'bold': True,'font_size':20,'fg_color': 'yellow'})
        information = workbook.add_format({'bold': True,'font_size':15})
        table_header = workbook.add_format({'bold': True,'font_size':12,'border':1 ,'font_color': 'white','fg_color': 'black'})
        
        table_row = workbook.add_format({'font_size':12})
        table_styled_row = workbook.add_format({'font_size':12,'fg_color': '#D9D9D9'})
        
        sheet5 = workbook.add_worksheet("Stock usine SPRINTER")
        sheet5_2 = workbook.add_worksheet("Listes des vins")

        sheet5.set_column('A:A', 30)
        sheet5.set_column('B:B', 30)
        sheet5.set_column('C:C', 30)
        sheet5.set_column('D:D', 30)
        sheet5.set_column('E:E', 30)
        sheet5.set_column('F:F', 30)
        sheet5.set_column('G:G', 30)
        sheet5.set_column('H:H', 30)
        sheet5.set_column('I:I', 30)
        sheet5.set_column('J:J', 30)
        sheet5.set_column('K:K', 30)
        sheet5.set_column('L:L', 30)
        sheet5.set_column('M:M', 30)
        sheet5.set_column('N:N', 30)
        sheet5.set_column('O:O', 30)
        sheet5.set_column('P:P', 30)

        sheet5_2.set_column('A:A', 30)
        sheet5_2.set_column('B:B', 30) 
        #sheet5_2.set_column('C:C', 40)          

        i = 1        
        sheet5.write(i, 2, "Stock usine SPRINTER du "+str(date_from.strftime("%d/%m/%Y")) , table_header)
        i = 3

        sheet5.write(i, 0, "Type2", table_header)  
        sheet5.write(i, 1, "Type3", table_header)
        sheet5.write(i, 2, "Variante Base", table_header)
        sheet5.write(i, 3, "Variante Commercial", table_header)
        sheet5.write(i, 4, "Stock K0", table_header)
        sheet5.write(i, 5, "En-Cours K1", table_header)
        sheet5.write(i, 6, "Stock K2", table_header)
        sheet5.write(i, 7, "En-Cours KC", table_header)
        sheet5.write(i, 8, "Stock K3", table_header)
        sheet5.write(i, 9, "Total Général", table_header)

        j = 1        
        sheet5_2.write(j, 2, "Stock usine SPRINTER du "+str(date_from.strftime("%d/%m/%Y")) , table_header)
        j = 3

        sheet5_2.write(j, 0, "Vin", table_header)  
        sheet5_2.write(j, 1, "Zone", table_header)
        sheet5_2.write(j, 2, "Variante Base", table_header)
        sheet5_2.write(j, 3, "Variante Carrosserie", table_header)
#***********26/09/24 add code and price
        sheet5_2.write(j, 4, "code article", table_header)
        sheet5_2.write(j, 5, "prix standard", table_header)

        col_1 = i + 1
        col_2 = i + 1
        col_3 = i + 1
        col_4 = i + 1
        col_5 = i + 1
        col_6 = i + 1

        for key in vin.keys():

            j += 1
            sheet5_2.write(j, 0, key, table_row)

            try:                                                           
                sheet5_2.write(j, 1, vin[key][0], table_row)
            except KeyError:
                pass

            try:                                                           
                sheet5_2.write(j, 2, vin[key][1], table_row)
            except KeyError:
                pass

            try:                                                           
                sheet5_2.write(j, 3, vin[key][2], table_row)
            except KeyError:
                pass
                
        
        for key_1 in structure.keys():                 
           
            for key_2 in structure[key_1].keys():                

                for key_3 in structure[key_1][key_2]:

                    for variant in structure[key_1][key_2][key_3]: 

                        sheet5.write(col_4, 3, variant, table_row)                  
                        col_d = col_4

                        for zone in ['k0','k1','k2','kc','k3','Total Général']:
                            
                            key_3_splited = key_3.split()
                            if len(key_3_splited) == 5:                            
                                color = key_3_splited[3]+' '+key_3_splited[4]
                            else:
                                color = key_3_splited[3]

                            if zone == 'Total Général':
                                if (variant == key_3):
                                    try:                                                           
                                        sheet5.write(col_d, 9, k['s'][key_3_splited[2]][color]['base'], table_header)
                                    except KeyError:
                                        pass
                                else:
                                    try:                                                           
                                        sheet5.write(col_d, 9, k['s'][key_3_splited[2]][color][variant], table_header)
                                    except KeyError:
                                        pass
                            else:
                                if (variant == key_3):                                   
                                    try:                                                                                                
                                        sheet5.write(col_d,index[zone], k[zone][key_3_splited[2]][color]['base'], table_row)
                                    except KeyError:
                                        pass
                                else:                                    
                                    try:                                                                                             
                                        sheet5.write(col_d, index[zone], k[zone][key_3_splited[2]][color][variant], table_row)
                                    except KeyError:
                                        pass           
                        
                        col_4 += 1
                    
                    if col_3 == col_4 - 1:
                        sheet5.write(col_3, 2, key_3, table_row)
                    else:
                        sheet5.merge_range(col_3, 2,col_4-1,2, key_3, table_row)

                    col_3 = col_4

                if key_2.startswith('Total'):
                    for zone in ['k0','k1','k2','kc','k3','Total Général']:
                        
                        try:    
                            sheet5.write(col_2, index[zone],s[key_2][zone] , table_header)
                        except KeyError:
                            pass

                    sheet5.merge_range(col_2, 1,col_2,3, key_2, table_header)
                    col_4 += 1                  
                    col_3 = col_4
                                 
                else:
                    if col_2 == col_4 - 1:
                        sheet5.write(col_2, 1, key_2, table_row)
                    else:
                        sheet5.merge_range(col_2, 1,col_4-1,1, key_2, table_row)

                col_2 = col_4
            
                
            if key_1.startswith('Total'):
                
                if key_1 == 'Total Général':                
                    for zone in ['k0','k1','k2','kc','k3','Total Général']:
                        try:    
                            sheet5.write(col_1, index[zone],st[zone] , table_header)
                        except KeyError:
                            pass
                    
                else:
                    for zone in ['k0','k1','k2','kc','k3','Total Général']:
                        
                        try:    
                            sheet5.write(col_1, index[zone],s[key_1][zone] , table_header)
                        except KeyError:
                            pass
                        

                               
                sheet5.merge_range(col_1, 0,col_1,3, key_1, table_header)
                col_4 += 1
                col_2 = col_4
                col_3 = col_4

                   
            else:
                sheet5.merge_range(col_1, 0,col_4-1,0, key_1, table_row)

            col_1 = col_4


        # ********************************** STOCK USINE ************************************


        # ********************************** PRODUCTION CARROSSERIE ************************************
     
        
        # ********************************** STOCK USINE ************************************


        # ********************************** PRODUCTION CARROSSERIE ************************************
        
class SCarrosserieProductionExcel(models.AbstractModel):
    _name = 'report.mb_reporting.sprinter_carrosserie_production_excel'
    _inherit = 'report.report_xlsx.abstract'
    
    def generate_xlsx_report(self, workbook, data, lines):
        
        cp = defaultdict()
        cp[5620] = 'A1 BUS 8+1'
        cp[5621] = 'A2 BUS 11+1 -VIP'
        cp[5622] = 'A2 BUS 14+1'
        cp[5623] = 'A2 BUS 14+1 4x4'
        cp[5624] = 'A4 BUS 20+1 -VIP'
        cp[5625] = 'A4 BUS 23+1'
        cp[5626] = 'A4 BUS 23+1 -Ecolier'
        cp[5627] = 'A4 BUS 23+1 4x4'
        cp[5629] = 'A4 BUS 29+1 -Ecolier'
        cp[5636] = 'A1 Cellulaire -Gendarmerie'
        cp[5635] = 'A1 BUS 8+1'
        cp[5637] = 'A1 Cellulaire -Police'
        cp[5638] = 'A1 VIP  Luxe 4+1'
        cp[5639] = 'A2 VIP  Luxe 6+2+1'
        cp[5646] = 'A4 Intervention -Gendarmerie'
        cp[5647] = 'A4 Intervention -Police'
        cp[5660] = 'A2 Ambulance Sanitaire'
        cp[5661] = 'A2 Ambulance Sanitaire 4x4'
        cp[5662] = 'A2 Ambulance Médicalisé'
        cp[5663] = 'A2 Ambulance Médicalisé 4x4'
        cp[5664] = 'A2 Ambulance Détenus'
        cp[5665] = 'A2 Ambulance Animalière'        
        cp[5685] = 'A1 Atelier Mobile'
        cp[5686] = 'A2 Atelier Mobile'
        cp[5695] = 'A2 Cabine Approfondie'
        cp[5696] = 'A1 Cabine Approfondie'
        cp[5705] = 'A2 Nacelle -12M'        
        cp[5706] = 'A2 Nacelle -14M'
        cp[5720] = 'A2 Frigo'
        cp[5721] = 'A2 Mortuaire'
        cp[5722] = 'A2 Transport De Chiens'
        cp[5723] = 'A2 Transport De Fonds'
        cp[5724] = 'A2 Intervention - Police'
        cp[5725] = 'A2 Transport De Détenus'

        A = defaultdict()
        


        k = defaultdict(lambda: defaultdict(lambda: defaultdict(dict)))
        variant1 = defaultdict()
        variant2 = defaultdict()
        vin = defaultdict(list)

        s = defaultdict(dict)
        st = defaultdict() 
        structure = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))

        a1t = defaultdict(list)
        a1v = defaultdict(list)    

        a2t = defaultdict(list)
        a2v = defaultdict(list)      

        a4t = defaultdict(list)
        a4v = defaultdict(list)      

        a1t = {'Fourgon Tolé S13 Blanc arctique':['A1 Atelier Mobile','A1 Cabine Approfondie']}
        a1v = {'Fourgon Vitré S02 Blanc arctique':['A1 BUS 8+1','A1 Cabine Approfondie','A1 Cellulaire -Police'],'Fourgon Vitré S12 Blanc arctique':['A1 BUS 8+1','A2 BUS 14+1'],'Fourgon Vitré S12 Gris Noir':['A1 BUS 8+1','A1 VIP Luxe 4+1']}

        a2t = {'Fourgon Tolé S04 Blanc arctique':['A2 Atelier Mobile','A2 Transport De Détenus'],'Fourgon Tolé S30 Blanc arctique':['A2 Nacelle -14M']}
        a2v = {'Fourgon Vitré S06 Blanc arctique':['A2 BUS 11+1 -VIP','A2 BUS 14+1'],'Fourgon Vitré S06 Gris Noir':['A2 BUS 11+1 -VIP','A2 BUS 14+1','A2 VIP  Luxe 6+2+1'],'Fourgon Vitré S10 Blanc arctique':['A2 Ambulance Animalière','A2 Ambulance Détenus','A2 Ambulance Médicalisé','A2 Ambulance Sanitaire','A2 Transport De Chiens','A2 Transport De Détenus']}
        
        a4t = {'Fourgon Tolé S25 Blanc arctique':['A4 Intervention -Gendarmerie']}
        a4v = {'Fourgon Vitré S09 Blanc arctique':['A4 BUS 20+1 -VIP','A4 BUS 23+1'],'Fourgon Vitré S09 Gris Noir':['A4 BUS 20+1 -VIP','A4 BUS 23+1'],'Fourgon Vitré S11 Jaune Genêt':['A4 BUS 23+1','A4 BUS 29+1 -Ecolier']}



        structure['A1'] = {'A1 Tolé':a1t,'Total A1 Tolé':{},'A1 Vitré':a1v,'Total A1 Vitré':{},}
        structure['Total A1'] = {}
        structure['A2'] = {'A2 Tolé':a2t,'Total A2 Tolé':{},'A2 Vitré':a2v,'Total A2 Vitré':{},}
        structure['Total A2'] = {}
        structure['A4'] = {'A4 Tolé':a4t,'Total A4 Tolé':{},'A4 Vitré':a4v,'Total A4 Vitré':{},}
        structure['Total A4'] = {}
        structure['Total Général'] = {}

        variant1['S13']='Total A1 Tolé'
        variant1['S02']='Total A1 Vitré'
        variant1['S12']='Total A1 Vitré'

        variant1['S04']='Total A2 Tolé'
        variant1['S14']='Total A2 Tolé'
        variant1['S30']='Total A2 Tolé'
        variant1['S06']='Total A2 Vitré'
        variant1['S10']='Total A2 Vitré'
        variant1['S22']='Total A2 Vitré'
        
        variant1['S23']='Total A4 Tolé'
        variant1['S25']='Total A4 Tolé'
        variant1['S27']='Total A4 Tolé'
        variant1['S09']='Total A4 Vitré'
        variant1['S11']='Total A4 Vitré'       

        variant2['S13']='Total A1'
        variant2['S02']='Total A1'
        variant2['S12']='Total A1'

        variant2['S04']='Total A2'
        variant2['S14']='Total A2'
        variant2['S30']='Total A2'
        variant2['S06']='Total A2'
        variant2['S10']='Total A2'
        variant2['S22']='Total A2'
        
        variant2['S23']='Total A4'
        variant2['S25']='Total A4'
        variant2['S27']='Total A4'
        variant2['S09']='Total A4'
        variant2['S11']='Total A4'      

        

        host = '192.168.1.34'
        database = 'TPR_Reports'
        user = 'safavmb\\odoo'
        password = 'Odoo@18'               

        date_from = data['form']['date_from']
        date_to = data['form']['date_to']

        date_from = datetime.strptime(date_from, '%Y-%m-%d')
        date_to = datetime.strptime(date_to, '%Y-%m-%d')

        DB = pymssql.connect(host=host,user=user,password=password,database=database)
        cursor = DB.cursor()
        
        cursor.callproc("[Reports].[spProductionCarrosserieSprinter]",[date_from,date_to])
        rows = list(cursor)
        for row in rows:
        
            if len(row[2]) == 12:
                if str(row[4]) not in ['UNKNOWN',None]:
                    color = str(row[4])
                vin[row[1]] = ["",""]
                vin[row[1]][0] = 'Fourgon ' + variant1[row[2][len(row[2])-3:]][5:] + ' (' + row[2][len(row[2])-3:] + ') ' + color
                vin[row[1]][1] = cp[row[5]]
                try:                
                    st[str(row[6].month)] += 1
                except KeyError:
                    st[str(row[6].month)] = 0
                    st[str(row[6].month)] += 1 
                
                try:
                    s[variant1[row[2][len(row[2])-3:]]][str(row[6].month)] += 1
                except KeyError:
                    s[variant1[row[2][len(row[2])-3:]]][str(row[6].month)] = 0
                    s[variant1[row[2][len(row[2])-3:]]][str(row[6].month)] += 1
                
                try:
                    s[variant2[row[2][len(row[2])-3:]]][str(row[6].month)] += 1
                except KeyError:
                    s[variant2[row[2][len(row[2])-3:]]][str(row[6].month)] = 0
                    s[variant2[row[2][len(row[2])-3:]]][str(row[6].month)] += 1

                try:                
                    st[str(13)] += 1
                except KeyError:
                    st[str(13)] = 0
                    st[str(13)] += 1 

                try:
                    s[variant1[row[2][len(row[2])-3:]]][str(13)] += 1
                except KeyError:
                    s[variant1[row[2][len(row[2])-3:]]][str(13)] = 0
                    s[variant1[row[2][len(row[2])-3:]]][str(13)] += 1
                
                try:
                    s[variant2[row[2][len(row[2])-3:]]][str(13)] += 1
                except KeyError:
                    s[variant2[row[2][len(row[2])-3:]]][str(13)] = 0
                    s[variant2[row[2][len(row[2])-3:]]][str(13)] += 1           

                try:                
                    k[row[2][len(row[2])-3:]][color][str(cp[row[5]])][str(row[6].month)] += 1
                except KeyError:
                    k[row[2][len(row[2])-3:]][color][str(cp[row[5]])][str(row[6].month)] = 0
                    k[row[2][len(row[2])-3:]][color][str(cp[row[5]])][str(row[6].month)] += 1

                try:                
                    k[row[2][len(row[2])-3:]][color][str(cp[row[5]])]['13'] += 1
                except KeyError:
                    k[row[2][len(row[2])-3:]][color][str(cp[row[5]])]['13'] = 0
                    k[row[2][len(row[2])-3:]][color][str(cp[row[5]])]['13'] += 1

                
        header = workbook.add_format({'bold': True,'font_size':20,'fg_color': 'yellow'})
        information = workbook.add_format({'bold': True,'font_size':15})
        table_header = workbook.add_format({'bold': True,'font_size':12,'border':1 ,'font_color': 'white','fg_color': 'black'})
        
        table_row = workbook.add_format({'font_size':12})
        table_styled_row = workbook.add_format({'font_size':12,'fg_color': '#D9D9D9'})
        
        sheet6 = workbook.add_worksheet("Production Carrosserie SPRINTER")
        sheet6_2 = workbook.add_worksheet("Liste Des Vins")

        sheet6.set_column('A:A', 15)
        sheet6.set_column('B:B', 15)
        sheet6.set_column('C:C', 30)
        sheet6.set_column('D:D', 30)
        sheet6.set_column('E:E', 10)
        sheet6.set_column('F:F', 10)
        sheet6.set_column('G:G', 10)
        sheet6.set_column('H:H', 10)
        sheet6.set_column('I:I', 10)
        sheet6.set_column('J:J', 10)
        sheet6.set_column('K:K', 10)
        sheet6.set_column('L:L', 10)
        sheet6.set_column('M:M', 10)
        sheet6.set_column('N:N', 10)
        sheet6.set_column('O:O', 10)        
        sheet6.set_column('P:P', 10)
        sheet6.set_column('Q:Q', 10)
        sheet6.set_column('R:R', 10)
        sheet6.set_column('S:S', 10)
        sheet6.set_column('T:T', 10)
        sheet6.set_column('U:U', 10)

        i = 1        
        sheet6.write(i, 2, "Production Carrosserie SPRINTER du "+str(date_from.strftime("%d/%m/%Y")) + " au " + str(date_to.strftime("%d/%m/%Y")), table_header)
        i = 3

        j = 1        
        sheet6_2.write(j, 2, "Production Carrosserie SPRINTER du "+str(date_from.strftime("%d/%m/%Y")) + " au " + str(date_to.strftime("%d/%m/%Y")) , table_header)
        j = 3
        
        sheet6_2.write(j, 0, "Vin", table_header)  
        sheet6_2.write(j, 1, "Variante Base", table_header)
        sheet6_2.write(j, 2, "Variante Carrosserie", table_header)

        sheet6.write(i, 0, "Type2", table_header)
        sheet6.write(i, 1, "Type3", table_header)
        sheet6.write(i, 2, "Variante Base", table_header)
        sheet6.write(i, 3, "Désignation Produit", table_header)
        sheet6.write(i, 4, "janv", table_header)
        sheet6.write(i, 5, "févr", table_header)
        sheet6.write(i, 6, "mars", table_header)
        sheet6.write(i, 7, "avr", table_header)
        sheet6.write(i, 8, "mai", table_header)
        sheet6.write(i, 9, "juin", table_header)
        sheet6.write(i, 10, "juil", table_header)
        sheet6.write(i, 11, "août", table_header)
        sheet6.write(i, 12, "sept", table_header)
        sheet6.write(i, 13, "oct", table_header)
        sheet6.write(i, 14, "nov", table_header)
        sheet6.write(i, 15, "déc", table_header)
        sheet6.write(i, 16, "Total Général", table_header)

        col_1 = i + 1
        col_2 = i + 1
        col_3 = i + 1
        col_4 = i + 1

        for key in vin.keys():
        
            j += 1
            sheet6_2.write(j, 0, key, table_row)

            try:                                                           
                sheet6_2.write(j, 1, vin[key][0], table_row)
            except KeyError:
                pass
            try:                                                           
                sheet6_2.write(j, 2, vin[key][1], table_row)
            except KeyError:
                pass

        for key_1 in structure.keys():                      
           
            for key_2 in structure[key_1].keys():                 

                for key_3 in structure[key_1][key_2]:

                    for variant in structure[key_1][key_2][key_3]: 

                        sheet6.write(col_4, 3, variant, table_row)                  
                        col_d = col_4

                        for l in (list (range(int(date_from.month),int(date_to.month)+1)) + [13,]):
                            
                            if (l==13):
                                format = table_header
                            else:
                                format = table_row
                            
                            try:
                                if key_3 != '':
                                    key_3_splited = key_3.split()
                                    if len(key_3_splited) == 5:                            
                                        sheet6.write(col_d, l+3, k[key_3_splited[2]][key_3_splited[3]+' '+key_3_splited[4]][variant][str(l)], format)
                                    else:
                                        sheet6.write(col_d, l+3, k[key_3_splited[2]][key_3_splited[3]][variant][str(l)], format)
                            except KeyError:
                                pass
                            
                        
                        col_4 += 1
                    

                    if col_3 == col_4 - 1:
                        sheet6.write(col_3, 2, key_3, table_row)
                    else:
                        sheet6.merge_range(col_3, 2,col_4-1,2, key_3, table_row)

                    col_3 = col_4

                if key_2.startswith('Total'):
                    for l in (list (range(int(date_from.month),int(date_to.month)+1)) + [13,]):
                        try:
                            sheet6.write(col_4, l+3, s[key_2][str(l)], table_header)
                        except KeyError:
                            pass       
                    sheet6.merge_range(col_2, 1,col_2,3, key_2, table_header)
                    col_4 += 1                  
                    col_3 = col_4
                                 
                else:
                    if col_2 == col_4 - 1:
                        sheet6.write(col_2, 1, key_2, table_row)
                    else:
                        sheet6.merge_range(col_2, 1,col_4-1,1, key_2, table_row)

                col_2 = col_4
            
                
            if key_1.startswith('Total'):                
                for l in (list (range(int(date_from.month),int(date_to.month)+1)) + [13,]):
                    if key_1 == 'Total Général':                                    
                        try:
                            sheet6.write(col_4, l+3, st[str(l)], table_header)
                        except KeyError:
                            pass 
                    else:
                        try:
                            sheet6.write(col_4, l+3, s[key_1][str(l)], table_header)
                        except KeyError:
                            pass  
                               
                sheet6.merge_range(col_1, 0,col_1,3, key_1, table_header)
                col_4 += 1
                col_2 = col_4
                col_3 = col_4

                   
            else:
                sheet6.merge_range(col_1, 0,col_4-1,0, key_1, table_row)

            col_1 = col_4

        


        # ********************************** PRODUCTION CARROSSERIE ************************************



        # ********************************** PRODUCTION BASE ************************************
        
class SBaseProductionExcel(models.AbstractModel):
    _name = 'report.mb_reporting.sprinter_base_production_excel'
    _inherit = 'report.report_xlsx.abstract'
    
    def generate_xlsx_report(self, workbook, data, lines):

        k = defaultdict(lambda: defaultdict(dict))               
        s = defaultdict(dict)
        st = defaultdict() 
        variant1 = defaultdict()
        variant2 = defaultdict()
        vin = defaultdict()



        structure = defaultdict(lambda: defaultdict(list))

        structure['A1'] = {'A1 Tolé':['Fourgon Tolé S13 Blanc arctique',],'Total A1 Tolé':[''],'A1 Vitré':['Fourgon Vitré S02 Blanc arctique','Fourgon Vitré S12 Blanc arctique','Fourgon Vitré S12 Gris Noir'],'Total A1 Vitré':[''],}
        structure['Total A1'] = {'empty':['']}
        structure['A2'] = {'A2 Tolé':['Fourgon Tolé S04 Blanc arctique','Fourgon Tolé S30 Blanc arctique','Fourgon Tolé S14 Noir',],'Total A2 Tolé':[''],'A2 Vitré':['Fourgon Vitré S06 Blanc arctique','Fourgon Vitré S06 Gris Noir','Fourgon Vitré S10 Blanc arctique','Fourgon Vitré S22 Blanc arctique',],'Total A2 Vitré':['']}
        structure['Total A2'] = {'empty':['']}
        structure['A4'] = {'A4 Tolé':['Fourgon Tolé S25 Blanc arctique','Fourgon Tolé S27 Blanc arctique',],'Total A4 Tolé':[''],'A4 Vitré':['Fourgon Vitré S09 Blanc arctique','Fourgon Vitré S09 Gris Noir'],'Total A4 Vitré':[''],}
        structure['Total A4'] = {'empty':['']}
        structure['Total Général'] = {'empty':['']}

        variant1['S13']='Total A1 Tolé'
        variant1['S02']='Total A1 Vitré'
        variant1['S12']='Total A1 Vitré'

        variant1['S04']='Total A2 Tolé'
        variant1['S14']='Total A2 Tolé'
        variant1['S30']='Total A2 Tolé'
        variant1['S06']='Total A2 Vitré'
        variant1['S10']='Total A2 Vitré'
        variant1['S22']='Total A2 Vitré'
        
        variant1['S23']='Total A4 Tolé'
        variant1['S25']='Total A4 Tolé'
        variant1['S27']='Total A4 Tolé'
        variant1['S09']='Total A4 Vitré'
        variant1['S11']='Total A4 Vitré'    

        

        variant2['S13']='Total A1'
        variant2['S02']='Total A1'
        variant2['S12']='Total A1'

        variant2['S04']='Total A2'
        variant2['S14']='Total A2'
        variant2['S30']='Total A2'
        variant2['S06']='Total A2'
        variant2['S10']='Total A2'
        variant2['S22']='Total A2'
        
        variant2['S23']='Total A4'
        variant2['S25']='Total A4'
        variant2['S27']='Total A4'
        variant2['S09']='Total A4'
        variant2['S11']='Total A4' 

        host = '192.168.1.34'
        database = 'TPR_Reports'
        user = 'safavmb\\odoo'
        password = 'Odoo@18'               

        date_from = data['form']['date_from']
        date_to = data['form']['date_to']

        date_from = datetime.strptime(date_from, '%Y-%m-%d')
        date_to = datetime.strptime(date_to, '%Y-%m-%d')

        DB = pymssql.connect(host=host,user=user,password=password,database=database)
        cursor = DB.cursor()
        
        cursor.callproc("[Reports].[spProductionBaseSprinter]",[date_from,date_to])
        rows = list(cursor)
        for row in rows:
        
            if len(row[2]) == 12:
                if str(row[4]) not in ['UNKNOWN',None]:
                    color = str(row[4])             

                vin[row[1]] = 'Fourgon ' + variant1[row[2][len(row[2])-3:]][5:] + ' (' + row[2][len(row[2])-3:] + ') ' + color             
              

                try:                
                    st[str(row[6].month)] += 1
                except KeyError:
                    st[str(row[6].month)] = 0
                    st[str(row[6].month)] += 1 
                
                try:
                    s[variant1[row[2][len(row[2])-3:]]][str(row[6].month)] += 1
                except KeyError:
                    s[variant1[row[2][len(row[2])-3:]]][str(row[6].month)] = 0
                    s[variant1[row[2][len(row[2])-3:]]][str(row[6].month)] += 1
                
                try:
                    s[variant2[row[2][len(row[2])-3:]]][str(row[6].month)] += 1
                except KeyError:
                    s[variant2[row[2][len(row[2])-3:]]][str(row[6].month)] = 0
                    s[variant2[row[2][len(row[2])-3:]]][str(row[6].month)] += 1

                try:                
                    st[str(13)] += 1
                except KeyError:
                    st[str(13)] = 0
                    st[str(13)] += 1 
                
                try:
                    s[variant1[row[2][len(row[2])-3:]]][str(13)] += 1
                except KeyError:
                    s[variant1[row[2][len(row[2])-3:]]][str(13)] = 0
                    s[variant1[row[2][len(row[2])-3:]]][str(13)] += 1
                
                try:
                    s[variant2[row[2][len(row[2])-3:]]][str(13)] += 1
                except KeyError:
                    s[variant2[row[2][len(row[2])-3:]]][str(13)] = 0
                    s[variant2[row[2][len(row[2])-3:]]][str(13)] += 1

                try:
                    k[row[2][len(row[2])-3:]][color][str(row[6].month)] += 1
                except KeyError:
                    k[row[2][len(row[2])-3:]][color][str(row[6].month)] = 0
                    k[row[2][len(row[2])-3:]][color][str(row[6].month)] += 1
                
                try:                
                    k[row[2][len(row[2])-3:]][color]['13'] += 1
                except KeyError:
                    k[row[2][len(row[2])-3:]][color]['13'] = 0
                    k[row[2][len(row[2])-3:]][color]['13'] += 1


        header = workbook.add_format({'bold': True,'font_size':20,'fg_color': 'yellow'})
        information = workbook.add_format({'bold': True,'font_size':15})
        table_header = workbook.add_format({'bold': True,'font_size':12,'border':1 ,'font_color': 'white','fg_color': 'black'})        
        table_row = workbook.add_format({'font_size':12})
        table_styled_row = workbook.add_format({'font_size':12,'fg_color': '#D9D9D9'})        
        sheet7 =workbook.add_worksheet("Production Base SPRINTER")
        sheet7_2 = workbook.add_worksheet("Listes des vins")

        sheet7.set_column('A:A', 20)
        sheet7.set_column('B:B', 30)
        sheet7.set_column('C:C', 30)
        sheet7.set_column('D:D', 10)
        sheet7.set_column('E:E', 10)
        sheet7.set_column('F:F', 10)
        sheet7.set_column('G:G', 10)
        sheet7.set_column('H:H', 10)
        sheet7.set_column('I:I', 10)
        sheet7.set_column('J:J', 10)
        sheet7.set_column('K:K', 10)
        sheet7.set_column('L:L', 10)
        sheet7.set_column('M:M', 10)
        sheet7.set_column('N:N', 10)
        sheet7.set_column('O:O', 10)        
        sheet7.set_column('P:P', 10)
        sheet7.set_column('Q:Q', 10)
        sheet7.set_column('R:R', 10)
        sheet7.set_column('S:S', 10)
        sheet7.set_column('T:T', 10)
        sheet7.set_column('U:U', 10)

        i = 1        
        sheet7.write(i, 2, "Production Base SPRINTER du "+str(date_from.strftime("%d/%m/%Y")) + " au " + str(date_to.strftime("%d/%m/%Y")), table_header)
        i = 3

        j = 1        
        sheet7_2.write(j, 2, "Production Base SPRINTER du "+str(date_from.strftime("%d/%m/%Y")) + " au " + str(date_to.strftime("%d/%m/%Y")) , table_header)
        j = 3
        
        sheet7_2.write(j, 0, "Vin", table_header)  
        sheet7_2.write(j, 1, "Variante Base", table_header)   
        

        sheet7.write(i, 0, "Type2", table_header)
        sheet7.write(i, 1, "Type3", table_header)
        sheet7.write(i, 2, "Variante Base", table_header)
        sheet7.write(i, 3, "janv", table_header)
        sheet7.write(i, 4, "févr", table_header)
        sheet7.write(i, 5, "mars", table_header)
        sheet7.write(i, 6, "avr", table_header)
        sheet7.write(i, 7, "mai", table_header)
        sheet7.write(i, 8, "juin", table_header)
        sheet7.write(i, 9, "juil", table_header)
        sheet7.write(i, 10, "août", table_header)
        sheet7.write(i, 11, "sept", table_header)
        sheet7.write(i, 12, "oct", table_header)
        sheet7.write(i, 13, "nov", table_header)
        sheet7.write(i, 14, "déc", table_header)
        sheet7.write(i, 15, "Total Général", table_header)
        
        col_1 = i + 1
        col_2 = i + 1
        col_3 = i + 1
        

        for key in vin.keys():
    
            j += 1
            sheet7_2.write(j, 0, key, table_row)

            try:                                                           
                sheet7_2.write(j, 1, vin[key], table_row)
            except KeyError:
                pass

           
        for key_1 in structure.keys():                      
           
            for key_2 in structure[key_1].keys():                 

                for variant in structure[key_1][key_2]: 

                    sheet7.write(col_3, 2, variant, table_row)                  
                    col_d = col_3

                    for l in (list (range(int(date_from.month),int(date_to.month)+1)) + [13,]):
                        
                        if (l==13):
                            format = table_header
                        else:
                            format = table_row

                        try:
                            if variant != '':
                                variant_splited = variant.split()
                                if len(variant_splited) == 5:                            
                                    sheet7.write(col_d, l+2, k[variant_splited[2]][variant_splited[3]+' '+variant_splited[4]][str(l)], format)
                                else:
                                    sheet7.write(col_d, l+2, k[variant_splited[2]][variant_splited[3]][str(l)], format)
                        except KeyError:
                            pass
                        if key_2.startswith('Total'):
                            try:
                                sheet7.write(col_d, l+2, s[key_2][str(l)], table_header)
                            except KeyError:
                                pass 
                        if key_1.startswith('Total'):
                           
                            if key_1 == 'Total Général':
                                
                                try:
                                    sheet7.write(col_d, l+2, st[str(l)], table_header)
                                except KeyError:
                                    pass 
                            else:
                                try:
                                    sheet7.write(col_d, l+2, s[key_1][str(l)], table_header)
                                except KeyError:
                                    pass                         
                    
                    col_3 += 1

                if key_2.startswith('Total'):
                    sheet7.merge_range(col_2, 1,col_2,2, key_2, table_header)                    
                else:
                    if col_2 == col_3 - 1:
                        sheet7.write(col_2, 1, key_2, table_row)
                    else:
                        sheet7.merge_range(col_2, 1,col_3-1,1, key_2, table_row)

                col_2 = col_3
            
             
            if key_1.startswith('Total'):
                sheet7.merge_range(col_1, 0,col_1,2, key_1, table_header)
                col_3 += 1
                col_2 = col_3
            else:
                sheet7.merge_range(col_1, 0,col_3-1,0, key_1, table_row)

            col_1 = col_3
                
                    

      
   
        # ********************************** PRODUCTION BASE ************************************

        # ************************************ SPRINTER *****************************************

        # ********************************** PRODUCTION BASE ************************************        

class GBaseProductionExcel(models.AbstractModel):
    _name = 'report.mb_reporting.gclass_base_production_excel'
    _inherit = 'report.report_xlsx.abstract'
    
    def generate_xlsx_report(self, workbook, data, lines):
        

        date_from = data['form']['date_from']
        date_to = data['form']['date_to']

        date_from = datetime.strptime(date_from, '%Y-%m-%d')
        date_to = datetime.strptime(date_to, '%Y-%m-%d')

        cp = defaultdict()
        ba = defaultdict(lambda: defaultdict(dict))
        s6 = defaultdict()
        s9 = defaultdict()
        s = defaultdict()
        vin = defaultdict()
        
        structure = defaultdict(list)
        variant_1 = defaultdict()

        variant_1['VLTT station long militaire Classe G 4X4 sable mat'] = 'm'
        variant_1['VLTT Station long militaire Classe G 4X4 vert bronze'] = 'm'
        variant_1['VLTT station long militaire Transmission (FFR) Classe G 4X4 sable mat'] = 'f'
        variant_1['VLTT station long militaire Transmission (FFR) Classe G 4X4 vert bronze'] = 'f'
        variant_1['VLTT Station long police avec rampe lumineuse Classe G 4X4 blanc polaire uni'] = 'p'
        
        variant_1['VLTT Chassis cabine civil Classe G 4X4 blanc polaire uni'] = 'c'
        variant_1['VLTT Chassis cabine civil Classe G 4X4 bleu tansanit métallique'] = 'c'  
        variant_1['VLTT Chassis cabine civil Classe G 4X4 rouge feu'] = 'c'
        variant_1['VLTT chassis cabine militaire Classe G 4X4 sable mat'] = 'm'  
        variant_1['VLTT Chassis cabine militaire Classe G 4X4 vert bronze'] = 'm'              

        structure['BA6'] = ['VLTT station long militaire Classe G 4X4 sable mat','VLTT Station long militaire Classe G 4X4 vert bronze','VLTT station long militaire Transmission (FFR) Classe G 4X4 sable mat','VLTT station long militaire Transmission (FFR) Classe G 4X4 vert bronze','VLTT Station long police avec rampe lumineuse Classe G 4X4 blanc polaire uni']
        structure['Total BA6'] = []
        structure['BA9'] = ['VLTT Chassis cabine civil Classe G 4X4 blanc polaire uni','VLTT Chassis cabine civil Classe G 4X4 bleu tansanit métallique','VLTT Chassis cabine civil Classe G 4X4 rouge feu','VLTT chassis cabine militaire Classe G 4X4 sable mat','VLTT Chassis cabine militaire Classe G 4X4 vert bronze']
        structure['Total BA9'] = []
        structure['Total Général'] = []       

        host = '192.168.1.34'
        database = 'TPR_Reports'
        user = 'safavmb\\odoo'
        password = 'Odoo@18'

        DB = pymssql.connect(host=host,user=user,password=password,database=database)
        cursor = DB.cursor()
        
        cursor.callproc("[Reports].[spProductionBaseGClass]",[date_from,date_to])
        rows = list(cursor)
        for row in rows:
           
            if str(row[1])[6:9]== '333':
                model = 'BA6'
                if  str(row[3]).startswith('BA 6 Police'):
                    model += 'p'
                    if row[4]!= None:
                        vin[row[1]] = 'VLTT Station long police avec rampe lumineuse Classe G 4X4 ' + row[4]

                if  str(row[3]).startswith('BA6 FFR'):
                    model += 'f'
                    if row[4]!= None:
                        vin[row[1]] = 'VLTT station long militaire Transmission (FFR) Classe G 4X4 ' + row[4]

                if  str(row[3]).startswith('BA 6 Mil'):
                    model += 'm'
                    if row[4]!= None:
                        vin[row[1]] = 'VLTT Station long militaire Classe G 4X4 ' + row[4]                
                
                try:                
                    s6[str(row[6].month)] += 1
                except KeyError:
                    s6[str(row[6].month)] = 0
                    s6[str(row[6].month)] += 1
                
                try:                
                    s6[str(13)] += 1
                except KeyError:
                    s6[str(13)] = 0
                    s6[str(13)] += 1
                
            if str(row[1])[6:9]== '343':
                model = 'BA9'
                if  str(row[3]).startswith('BA 9 Civil'):
                    model += 'c'
                    if row[4]!= None:
                        vin[row[1]] = 'VLTT Chassis cabine civil Classe G 4X4 ' + row[4]

                if  str(row[3]).startswith('BA 9 Military'):
                    model += 'm'
                    if row[4]!= None:
                        vin[row[1]] = 'VLTT Chassis cabine militaire Classe G 4X4 ' + row[4]

                try:                
                    s9[str(row[6].month)] += 1
                except KeyError:
                    s9[str(row[6].month)] = 0
                    s9[str(row[6].month)] += 1
                try:                
                    s9[str(13)] += 1
                except KeyError:
                    s9[str(13)] = 0
                    s9[str(13)] += 1

            try:                
                s[str(row[6].month)] += 1
            except KeyError:
                s[str(row[6].month)] = 0
                s[str(row[6].month)] += 1 
            
            try:                
                s[str(13)] += 1
            except KeyError:
                s[str(13)] = 0
                s[str(13)] += 1 

            try:                
                ba[model][str(row[4])][str(row[6].month)] += 1
            except KeyError:
                ba[model][str(row[4])][str(row[6].month)] = 0
                ba[model][str(row[4])][str(row[6].month)] += 1
            
            try:                
                ba[model][str(row[4])][str(13)] += 1
            except KeyError:
                ba[model][str(row[4])][str(13)] = 0
                ba[model][str(row[4])][str(13)] += 1

        header = workbook.add_format({'bold': True,'font_size':20,'fg_color': 'yellow'})
        information = workbook.add_format({'bold': True,'font_size':15})
        table_header = workbook.add_format({'bold': True,'font_size':12,'border':1 ,'font_color': 'white','fg_color': 'black'})
        
        table_row = workbook.add_format({'font_size':12})
        table_styled_row = workbook.add_format({'font_size':12,'fg_color': '#D9D9D9'})
        
        sheet8 = workbook.add_worksheet("Production Base G-CLASS")
        sheet8_2 = workbook.add_worksheet("Liste Des Vins")

        sheet8.set_column('A:A', 20)
        sheet8.set_column('B:B', 60)
        sheet8.set_column('C:C', 10)
        sheet8.set_column('D:D', 10)
        sheet8.set_column('E:E', 10)
        sheet8.set_column('F:F', 10)
        sheet8.set_column('G:G', 10)
        sheet8.set_column('H:H', 10)
        sheet8.set_column('I:I', 10)
        sheet8.set_column('J:J', 10)
        sheet8.set_column('K:K', 10)
        sheet8.set_column('L:L', 10)
        sheet8.set_column('M:M', 10)
        sheet8.set_column('N:N', 10)
        sheet8.set_column('O:O', 10)        
        sheet8.set_column('P:P', 10)
        sheet8.set_column('Q:Q', 10)
        sheet8.set_column('R:R', 10)
        sheet8.set_column('S:S', 10)
        sheet8.set_column('T:T', 10)
        sheet8.set_column('U:U', 10)

        i = 1        
        sheet8.write(i, 2, "Production Base G-CLASS du "+str(date_from.strftime("%d/%m/%Y")) + " au " + str(date_to.strftime("%d/%m/%Y")), table_header)
        i = 3

        j = 1        
        sheet8_2.write(j, 2, "Production Base G-CLASS du "+str(date_from.strftime("%d/%m/%Y")) + " au " + str(date_to.strftime("%d/%m/%Y")) , table_header)
        j = 3
        
        sheet8_2.write(j, 0, "Vin", table_header)  
        sheet8_2.write(j, 1, "Variante Base", table_header) 

        sheet8.write(i, 0, "Type", table_header)
        sheet8.write(i, 1, "Variante Base", table_header)
        sheet8.write(i, 2, "janv", table_header)
        sheet8.write(i, 3, "févr", table_header)
        sheet8.write(i, 4, "mars", table_header)
        sheet8.write(i, 5, "avr", table_header)
        sheet8.write(i, 6, "mai", table_header)
        sheet8.write(i, 7, "juin", table_header)
        sheet8.write(i, 8, "juil", table_header)
        sheet8.write(i, 9, "août", table_header)
        sheet8.write(i, 10, "sept", table_header)
        sheet8.write(i, 11, "oct", table_header)
        sheet8.write(i, 12, "nov", table_header)
        sheet8.write(i, 13, "déc", table_header)
        sheet8.write(i, 14, "Total Général", table_header)

        col_1 = i + 1
        col_2 = i + 1       

        for key in vin.keys():
        
            j += 1
            sheet8_2.write(j, 0, key, table_row)

            try:                                                           
                sheet8_2.write(j, 1, vin[key], table_row)
            except KeyError:
                pass


        for key_1 in structure.keys():                      
           
            for variant in structure[key_1]: 

                sheet8.write(col_2, 1, variant, table_row)                  
                col_d = col_2

                for l in (list (range(int(date_from.month),int(date_to.month)+1)) + [13,]):
                    
                    if (l==13):
                        format = table_header
                    else:
                        format = table_row
                    
                    try:
                        if key_1 != '':
                            variant_splited = variant.split()
                            if variant_splited[len(variant_splited)-1] in ['métallique','uni']:
                                color = variant_splited[len(variant_splited)-3]+' '+variant_splited[len(variant_splited)-2]+' '+variant_splited[len(variant_splited)-1]
                            else:
                                color = variant_splited[len(variant_splited)-2]+' '+variant_splited[len(variant_splited)-1]
                           
                           
                            sheet8.write(col_d, l+1, ba[str(key_1)+str(variant_1[variant])][color][str(l)], format)
                    except KeyError:
                        pass
                        
                col_2 += 1

                
            if key_1.startswith('Total'):                
                for l in (list (range(int(date_from.month),int(date_to.month)+1)) + [13,]):
                    if key_1 == 'Total Général':                                    
                        try:
                            sheet8.write(col_2, l+1, s[str(l)], table_header)
                        except KeyError:
                            pass 
                    
                    if key_1 == 'Total BA6':
                        try:
                            sheet8.write(col_2, l+1, s6[str(l)], table_header)
                        except KeyError:
                            pass  
                    
                    if key_1 == 'Total BA9':
                        try:
                            sheet8.write(col_2, l+1, s9[str(l)], table_header)
                        except KeyError:
                            pass  
                               
                sheet8.merge_range(col_1, 0,col_1,1, key_1, table_header)
                col_2 += 1
                

                   
            else:
                sheet8.merge_range(col_1, 0,col_2-1,0, key_1, table_row)

            col_1 = col_2

        # sheet8.merge_range(i+13, 0, i+13, 1, 'Total Général', table_header)
        # somme = sum(s[key] for key in s.keys() if key != '13')
        # sheet8.write(i+13, 14, somme, table_header)

        # ********************************** PRODUCTION BASE ************************************


        # ********************************** PRODUCTION CARROSSERIE ************************************        

class GBaseCarrosserieExcel(models.AbstractModel):
    _name = 'report.mb_reporting.gclass_carrosserie_production_excel'
    _inherit = 'report.report_xlsx.abstract'
    
    def generate_xlsx_report(self, workbook, data, lines):


        date_from = data['form']['date_from']
        date_to = data['form']['date_to']

        date_from = datetime.strptime(date_from, '%Y-%m-%d')
        date_to = datetime.strptime(date_to, '%Y-%m-%d')
        
        cp = defaultdict()
        vin = defaultdict(list)

        cp[5400]='Ambulance-CIVIL'
        cp[5410]='AMP'
        cp[5420]='Ambulance-Militaire'
        cp[5425]='ATM'
        cp[5430]='Anti-incendie'
        cp[5440]='Transport de troupes'
        cp[5445]='Transport de marchandises'
        cp[5450]='Giraffe Lumineuse'
        cp[5460]='FFR'
        cp[5470]='FFR Light'
        cp[5480]='DGPC'
        cp[5490]='LUX'
        cp[5500]='POL'
        cp[5610]='GEN'
        cp[5620]='DOU'
        
        ba = defaultdict(lambda: defaultdict(dict))
        s6 = defaultdict()
        s9 = defaultdict()
        s = defaultdict()        

        structure = defaultdict()       

        structure['BA6'] = ['FFR Light']
        structure['Total BA6'] = []
        structure['BA9'] = ['Ambulance-CIVIL','Ambulance-Militaire','FFR','Giraffe Lumineuse','Transport de marchandises','Transport de troupes']
        structure['Total BA9'] = []
        structure['Total Général'] = []       

        host = '192.168.1.34'
        database = 'TPR_Reports'
        user = 'safavmb\\odoo'
        password = 'Odoo@18'

        DB = pymssql.connect(host=host,user=user,password=password,database=database)
        cursor = DB.cursor()
        
        cursor.callproc("[Reports].[spProductionCarrosserieGClass]",[date_from,date_to])
        rows = list(cursor)
        for row in rows:

            
            if str(row[1])[6:9]== '333':
                model = 'BA6'
                try:                
                    s6[str(row[6].month)] += 1
                except KeyError:
                    s6[str(row[6].month)] = 0
                    s6[str(row[6].month)] += 1 
                try:                
                    s6[str(13)] += 1
                except KeyError:
                    s6[str(13)] = 0
                    s6[str(13)] += 1             
                
            if str(row[1])[6:9]== '343':
                model = 'BA9'
                try:                
                    s9[str(row[6].month)] += 1
                except KeyError:
                    s9[str(row[6].month)] = 0
                    s9[str(row[6].month)] += 1
                try:                
                    s9[str(13)] += 1
                except KeyError:
                    s9[str(13)] = 0
                    s9[str(13)] += 1
            try:                
                s[str(row[6].month)] += 1
            except KeyError:
                s[str(row[6].month)] = 0
                s[str(row[6].month)] += 1
            try:                
                s[str(13)] += 1
            except KeyError:
                s[str(13)] = 0
                s[str(13)] += 1

            vin[row[1]] = str(cp[row[5]])

            try:                
                ba[model][str(cp[row[5]])][str(row[6].month)] += 1
            except KeyError:
                ba[model][str(cp[row[5]])][str(row[6].month)] = 0
                ba[model][str(cp[row[5]])][str(row[6].month)] += 1
            
            try:                
                ba[model][str(cp[row[5]])][str(13)] += 1
            except KeyError:
                ba[model][str(cp[row[5]])][str(13)] = 0
                ba[model][str(cp[row[5]])][str(13)] += 1

        

        header = workbook.add_format({'bold': True,'font_size':20,'fg_color': 'yellow'})
        information = workbook.add_format({'bold': True,'font_size':15})
        table_header = workbook.add_format({'bold': True,'font_size':12,'border':1 ,'font_color': 'white','fg_color': 'black'})
        
        table_row = workbook.add_format({'font_size':12})
        table_styled_row = workbook.add_format({'font_size':12,'fg_color': '#D9D9D9'})

        
        sheet9 = workbook.add_worksheet("Production Carrosserie G-CLASS")
        sheet9_2 = workbook.add_worksheet("Liste Des Vins")

        sheet9.set_column('A:A', 20)
        sheet9.set_column('B:B', 30)
        sheet9.set_column('C:C', 10)
        sheet9.set_column('D:D', 10)
        sheet9.set_column('E:E', 10)
        sheet9.set_column('F:F', 10)
        sheet9.set_column('G:G', 10)
        sheet9.set_column('H:H', 10)
        sheet9.set_column('I:I', 10)
        sheet9.set_column('J:J', 10)
        sheet9.set_column('K:K', 10)
        sheet9.set_column('L:L', 10)
        sheet9.set_column('M:M', 10)
        sheet9.set_column('N:N', 10)
        sheet9.set_column('O:O', 10)        
        sheet9.set_column('P:P', 10)
        sheet9.set_column('Q:Q', 10)
        sheet9.set_column('R:R', 10)
        sheet9.set_column('S:S', 10)
        sheet9.set_column('T:T', 10)
        sheet9.set_column('U:U', 10)

        i = 1        
        sheet9.write(i, 2, "Production Carrosserie G-CLASS du "+str(date_from.strftime("%d/%m/%Y")) + " au " + str(date_to.strftime("%d/%m/%Y")), table_header)
        i = 3

        j = 1        
        sheet9_2.write(j, 2, "Production Carrosserie G-CLASS du "+str(date_from.strftime("%d/%m/%Y")) + " au " + str(date_to.strftime("%d/%m/%Y")) , table_header)
        j = 3
        
        sheet9_2.write(j, 0, "Vin", table_header)  
        sheet9_2.write(j, 1, "Variante Carrosserie", table_header) 

        sheet9.write(i, 0, "Type", table_header)
        sheet9.write(i, 1, "Désignation produit", table_header)
        sheet9.write(i, 2, "janv", table_header)
        sheet9.write(i, 3, "févr", table_header)
        sheet9.write(i, 4, "mars", table_header)
        sheet9.write(i, 5, "avr", table_header)
        sheet9.write(i, 6, "mai", table_header)
        sheet9.write(i, 7, "juin", table_header)
        sheet9.write(i, 8, "juil", table_header)
        sheet9.write(i, 9, "août", table_header)
        sheet9.write(i, 10, "sept", table_header)
        sheet9.write(i, 11, "oct", table_header)
        sheet9.write(i, 12, "nov", table_header)
        sheet9.write(i, 13, "déc", table_header)
        sheet9.write(i, 14, "Total Général", table_header)

        col_1 = i + 1
        col_2 = i + 1       
        for key in vin.keys():
            
            j += 1
            sheet9_2.write(j, 0, key, table_row)

            try:                                                           
                sheet9_2.write(j, 1, vin[key], table_row)
            except KeyError:
                pass

        for key_1 in structure.keys():                      
           
            for variant in structure[key_1]: 

                sheet9.write(col_2, 1, variant, table_row)                  
                col_d = col_2

                for l in (list (range(int(date_from.month),int(date_to.month)+1)) + [13,]):
                    
                    if (l==13):
                        format = table_header
                    else:
                        format = table_row
                    
                    try:
                        if key_1 != '':                            
                            sheet9.write(col_d, l+1, ba[key_1][variant][str(l)], format)
                    except KeyError:
                        pass
                        
                col_2 += 1

            if key_1.startswith('Total'):                
                for l in (list (range(int(date_from.month),int(date_to.month)+1)) + [13,]):
                    if key_1 == 'Total Général':                                    
                        try:
                            sheet9.write(col_2, l+1, s[str(l)], table_header)
                        except KeyError:
                            pass 
                    
                    if key_1 == 'Total BA6':
                        try:
                            sheet9.write(col_2, l+1, s6[str(l)], table_header)
                        except KeyError:
                            pass  
                    
                    if key_1 == 'Total BA9':
                        try:
                            sheet9.write(col_2, l+1, s9[str(l)], table_header)
                        except KeyError:
                            pass  
                               
                sheet9.merge_range(col_1, 0,col_1,1, key_1, table_header)
                col_2 += 1
                

                   
            else:
                sheet9.merge_range(col_1, 0,col_2-1,0, key_1, table_row)

            col_1 = col_2

        
        # ********************************** PRODUCTION CARROSSERIE ************************************

        # ********************************** STOCK USINE ************************************        

class GStockUsineExcel(models.AbstractModel):


    _name = 'report.mb_reporting.gclass_stock_usine_excel'
    _inherit = 'report.report_xlsx.abstract'
    
    def generate_xlsx_report(self, workbook, data, lines):        

        date_from = data['form']['date_from']
        date_from = datetime.strptime(date_from, '%Y-%m-%d')
        
        cp = defaultdict()

        cp[5200]='Ambulance-CIVIL'
        cp[5210]='AMP'
        cp[5220]='Ambulance-Militaire'
        cp[5225]='ATM'
        cp[5230]='Anti-incendie'
        cp[5240]='Transport de troupes'
        cp[5245]='Transport de marchandises'
        cp[5250]='Giraffe Lumineuse'
        cp[5260]='FFR'
        cp[5270]='FFR Light'
        cp[5280]='DGPC'
        cp[5290]='LUX'
        cp[5300]='POL'
        cp[5310]='GEN'
        cp[5320]='DOU'

        cp[5400]='Ambulance-CIVIL'
        cp[5410]='AMP'
        cp[5420]='Ambulance-Militaire'
        cp[5425]='ATM'
        cp[5430]='Anti-incendie'
        cp[5440]='Transport de troupes'
        cp[5445]='Transport de marchandises'
        cp[5450]='Giraffe Lumineuse'
        cp[5460]='FFR'
        cp[5470]='FFR Light'
        cp[5480]='DGPC'
        cp[5490]='LUX'
        cp[5500]='POL'
        cp[5610]='GEN'
        cp[5620]='DOU'

        k = defaultdict(lambda: defaultdict(lambda: defaultdict(dict)))        
        s6 = defaultdict()
        s9 = defaultdict()
        s = defaultdict()
        vin = defaultdict(list)

        variant_1 = defaultdict()

        variant_1['VLTT station long militaire Classe G 4X4 sable mat'] = 'm'
        variant_1['VLTT Station long militaire Classe G 4X4 vert bronze'] = 'm'
        variant_1['VLTT station long militaire Transmission (FFR) Classe G 4X4 sable mat'] = 'f'
        variant_1['VLTT station long militaire Transmission (FFR) Classe G 4X4 vert bronze'] = 'f'
        variant_1['VLTT Station long police avec rampe lumineuse Classe G 4X4 blanc polaire uni'] = 'p'
        
        variant_1['VLTT Chassis cabine civil Classe G 4X4 blanc polaire uni'] = 'c'
        variant_1['VLTT Chassis cabine civil Classe G 4X4 bleu tansanit métallique'] = 'c'  
        variant_1['VLTT Chassis cabine civil Classe G 4X4 rouge feu'] = 'c'
        variant_1['VLTT chassis cabine militaire Classe G 4X4 sable mat'] = 'm'  
        variant_1['VLTT Chassis cabine militaire Classe G 4X4 vert bronze'] = 'm' 
        variant_1['VLTT Chassis cabine militaire Transmission (FFR) Classe G 4X4 vert bronze'] = 'f'           

        structure = defaultdict(lambda: defaultdict(list))

        structure['BA6'] = {
                                'VLTT station long militaire Transmission (FFR) Classe G 4X4 sable mat':['VLTT station long militaire Transmission (FFR) Classe G 4X4 sable mat','FFR Light'],
                                'VLTT station long militaire Transmission (FFR) Classe G 4X4 vert bronze':['VLTT station long militaire Transmission (FFR) Classe G 4X4 vert bronze','FFR Light'],
                                'VLTT Station long police avec rampe lumineuse Classe G 4X4 blanc polaire uni':['VLTT Station long police avec rampe lumineuse Classe G 4X4 blanc polaire uni']
                            }
        
        structure['Total BA6'] = {}
        structure['BA9'] = {
                                'VLTT Chassis cabine civil Classe G 4X4 blanc polaire uni':['VLTT Chassis cabine civil Classe G 4X4 blanc polaire uni','Transport de marchandises'],
                                'VLTT Chassis cabine civil Classe G 4X4 bleu tansanit métallique':['VLTT Chassis cabine civil Classe G 4X4 bleu tansanit métallique','Giraffe Lumineuse'],
                                'VLTT Chassis cabine civil Classe G 4X4 rouge feu':['VLTT Chassis cabine civil Classe G 4X4 rouge feu','Anti-incendie'],
                                'VLTT chassis cabine militaire Classe G 4X4 Sable mat':['VLTT chassis cabine militaire Classe G 4X4 Sable mat'],
                                'VLTT Chassis cabine militaire Classe G 4X4 vert bronze':['VLTT Chassis cabine militaire Classe G 4X4 vert bronze','Ambulance-Militaire','Transport de troupes'],
                                'VLTT Chassis cabine militaire Transmission (FFR) Classe G 4X4 vert bronze':['VLTT Chassis cabine militaire Transmission (FFR) Classe G 4X4 vert bronze']
                            }
        
        structure['Total BA9'] = {}
        structure['Total Général'] = {}
        
        index = defaultdict()

        index['k0'] = 3
        index['k1'] = 4
        index['k2'] = 5
        index['kc'] = 6
        index['k3'] = 7

        host = '192.168.1.34'
        database = 'TPR_Reports'
        user = 'safavmb\\odoo'
        password = 'Odoo@18'

        DB = pymssql.connect(host=host,user=user,password=password,database=database)
        cursor = DB.cursor()      
       
        cursor.callproc("[Reports].[spStockUsineGClass]",[date_from + timedelta(days=1),])
        rows = list(cursor)
        for row in rows:
            
            if row[11] == 'k1-kc':
                if row[9]:
                    area = 'kc'
                else:
                    area = 'k1'
            else:
                area = row[11]

            vin[row[1]] = [area,"",""]

            if str(row[1])[6:9]== '333':
                model = 'BA6'
                if  str(row[3]).startswith('BA 6 Police'):
                    model += 'p'
                    if row[4]!= None:
                        vin[row[1]][1] = 'VLTT Station long police avec rampe lumineuse Classe G 4X4 ' + row[4]

                if  str(row[3]).startswith('BA6 FFR'):
                    model += 'f'
                    if row[4]!= None:
                        vin[row[1]][1] = 'VLTT station long militaire Transmission (FFR) Classe G 4X4 ' + row[4]

                if  str(row[3]).startswith('BA 6 Mil'):
                    model += 'm'
                    if row[4]!= None:
                        vin[row[1]][1] = 'VLTT station long militaire Classe G 4X4 ' + row[4]
                    
                try:                
                    s6[area] += 1
                except KeyError:
                    s6[area] = 0
                    s6[area] += 1

            if str(row[1])[6:9]== '343':
                model = 'BA9'
                if  str(row[3]).startswith('BA 9 Civil'):
                    model += 'c'
                    if row[4]!= None:
                        vin[row[1]][1] = 'VLTT Chassis cabine civil Classe G 4X4 ' + row[4]
                   

                if  str(row[3]).startswith('BA 9 Military'):
                    model += 'm'
                    if row[4]!= None:
                        vin[row[1]][1] = 'VLTT Chassis cabine militaire Classe G 4X4 ' + row[4]

                try:                
                    s9[area] += 1
                except KeyError:
                    s9[area] = 0
                    s9[area] += 1              
            
            if row[9] != None:
                vin[row[1]][2] = cp[row[9]]
                try:
                    k[area][model][str(row[4])][cp[row[9]]] += 1
                except KeyError:
                    k[area][model][str(row[4])][cp[row[9]]] = 0
                    k[area][model][str(row[4])][cp[row[9]]] += 1

                if row[11] != 'delivered to customer':
                    try:
                        k['s'][model][str(row[4])][cp[row[9]]] += 1
                    except KeyError:
                        k['s'][model][str(row[4])][cp[row[9]]] = 0
                        k['s'][model][str(row[4])][cp[row[9]]] += 1
            else:
                try:
                    k[area][model][str(row[4])]['base'] += 1
                except KeyError:
                    k[area][model][str(row[4])]['base'] = 0
                    k[area][model][str(row[4])]['base'] += 1
                
                if row[11] != 'delivered to customer':
                    try:
                        k['s'][model][str(row[4])]['base'] += 1
                    except KeyError:
                        k['s'][model][str(row[4])]['base'] = 0
                        k['s'][model][str(row[4])]['base'] += 1
            
            if row[11] != 'delivered to customer':
                try:                
                    s[area] += 1
                except KeyError:
                    s[area] = 0
                    s[area] += 1
        
        header = workbook.add_format({'bold': True,'font_size':20,'fg_color': 'yellow'})
        information = workbook.add_format({'bold': True,'font_size':15})
        table_header = workbook.add_format({'bold': True,'font_size':12,'border':1 ,'font_color': 'white','fg_color': 'black'})
        
        table_row = workbook.add_format({'font_size':12})
        table_styled_row = workbook.add_format({'font_size':12,'fg_color': '#D9D9D9'})

        
        sheet10 = workbook.add_worksheet("Stock Usine G-CLASS")
        sheet10_2 = workbook.add_worksheet("Listes Des Vins")

        sheet10.set_column('A:A', 30)
        sheet10.set_column('B:B', 30)
        sheet10.set_column('C:C', 30)
        sheet10.set_column('D:D', 30)
        sheet10.set_column('E:E', 30)
        sheet10.set_column('F:F', 30)
        sheet10.set_column('G:G', 30)
        sheet10.set_column('H:H', 30)
        sheet10.set_column('I:I', 30)
        sheet10.set_column('J:J', 30)
        sheet10.set_column('K:K', 30)
        sheet10.set_column('L:L', 30)
        sheet10.set_column('M:M', 30)
        sheet10.set_column('N:N', 30)
        sheet10.set_column('O:O', 30)        
        sheet10.set_column('P:P', 30)
        sheet10.set_column('Q:Q', 30)
        sheet10.set_column('R:R', 30)
        sheet10.set_column('S:S', 30)
        sheet10.set_column('T:T', 30)
        sheet10.set_column('U:U', 30)

        sheet10_2.set_column('A:A', 30)
        sheet10_2.set_column('B:B', 30)

        i = 1        
        sheet10.write(i, 2, "Stock Usine G-CLASS du "+str(date_from.strftime("%d/%m/%Y")) , table_header)
        i = 3

        j = 1        
        sheet10_2.write(j, 2, "Stock Usine G-CLASS du "+str(date_from.strftime("%d/%m/%Y")) , table_header)
        j = 3

        sheet10.write(i, 0, "Type", table_header)
        sheet10.write(i, 1, "Variante Base", table_header)
        sheet10.write(i, 2, "Désignation produit", table_header)
        sheet10.write(i, 3, "Stock K0", table_header)
        sheet10.write(i, 4, "En-Cours K1", table_header)
        sheet10.write(i, 5, "Stock K2", table_header)
        sheet10.write(i, 6, "En-Cours KC", table_header)
        sheet10.write(i, 7, "Stock K3", table_header)
        sheet10.write(i, 8, "Total Général", table_header)

        sheet10_2.write(j, 0, "Vin", table_header)  
        sheet10_2.write(j, 1, "Zone", table_header)
        sheet10_2.write(j, 2, "Variante Base", table_header)
        sheet10_2.write(j, 3, "Variante Carrosserie", table_header)

        col_1 = i + 1
        col_2 = i + 1      
        col_3 = i + 1  

        for key in vin.keys():
            j += 1
            sheet10_2.write(j, 0, key, table_row)
            sheet10_2.write(j, 1, vin[key][0], table_row)
            sheet10_2.write(j, 2, vin[key][1], table_row)
            sheet10_2.write(j, 3, vin[key][2], table_row)

        for key_1 in structure.keys(): 

            for key_2 in structure[key_1].keys():

                for variant in structure[key_1][key_2]: 

                    sheet10.write(col_3, 2, variant, table_row)                  
                    col_d = col_3

                    for zone in ['k0','k1','k2','kc','k3','Total Général']:

                        key_2_splited = key_2.split()
                        if key_2_splited[len(key_2_splited)-1] in ['métallique','uni']:
                            color = key_2_splited[len(key_2_splited)-3]+' '+key_2_splited[len(key_2_splited)-2]+' '+key_2_splited[len(key_2_splited)-1]
                        else:
                            color = key_2_splited[len(key_2_splited)-2]+' '+key_2_splited[len(key_2_splited)-1]
                        
                        if zone == 'Total Général':
                            if (variant == key_2):
                                try:                                                           
                                    sheet10.write(col_d, 8, k['s'][str(key_1)+variant_1[key_2]][color]['base'], table_header)
                                except KeyError:
                                    pass
                            else:
                                try:                                                           
                                    sheet10.write(col_d, 8, k['s'][str(key_1)+variant_1[key_2]][color][variant], table_header)
                                except KeyError:
                                    pass
                        else:
                            if (variant == key_2):
                                try:                                                           
                                    sheet10.write(col_d,index[zone], k[zone][str(key_1)+variant_1[key_2]][color]['base'], table_row)
                                except KeyError:
                                    pass
                            else:
                                try:                                                           
                                    sheet10.write(col_d, index[zone], k[zone][str(key_1)+variant_1[key_2]][color][variant], table_row)
                                except KeyError:
                                    pass                       
                        
                    
                    col_3 += 1

                if col_2 == col_3 - 1:
                    sheet10.write(col_2, 1, key_2, table_row)
                else:
                    sheet10.merge_range(col_2, 1,col_3-1,1, key_2, table_row)

                col_2 = col_3

            if key_1.startswith('Total'):
                
                somme = 0
                somme_6 = 0
                somme_9 = 0 

                for zone in ['k0','k1','k2','kc','k3','Total Général']:
                    
                    if key_1 == 'Total Général':
                                                            
                        try:
                            sheet10.write(col_3, index[zone], s[zone], table_header)
                            somme += s[zone] 
                        except KeyError:
                            pass

                        
                    
                    if key_1 == 'Total BA6':
                        try:
                            sheet10.write(col_3, index[zone], s6[zone], table_header)
                            somme_6 += s6[zone]
                        except KeyError:
                            pass  

                        

                    if key_1 == 'Total BA9':
                        try:
                            sheet10.write(col_3, index[zone], s9[zone], table_header)
                            somme_9 += s9[zone]
                        except KeyError:
                            pass

                if key_1 == 'Total Général':
                    sheet10.write(col_3, 8, somme, table_header)
                if key_1 == 'Total BA6':
                    sheet10.write(col_3, 8, somme_6, table_header)
                if key_1 == 'Total BA9':
                    sheet10.write(col_3, 8, somme_9, table_header)
                
               
                sheet10.merge_range(col_1, 0,col_1,2, key_1, table_header)
                col_3 += 1
                col_2 = col_3

                   
            else:
                sheet10.merge_range(col_1, 0,col_2-1,0, key_1, table_row)

            col_1 = col_3
        

        
        somme = sum(s[key] for key in s.keys())
        sheet10.write(col_3-1, 8, somme, table_header)
        
        # ********************************** STOCK USINE ************************************

        # ********************************* G-CLASS ********************************************


        # ********************************** STOCK USINE ************************************

        # ********************************* G-CLASS *********************************************
