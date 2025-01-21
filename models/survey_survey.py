from odoo import models, fields, api
from odoo.exceptions import UserError
import io
import xlsxwriter
import base64

class SurveySurvey(models.Model):
    _inherit = 'survey.survey'

    def export_survey_results(self):
        """
        Exporta los resultados de una encuesta a un archivo Excel.
        """
        self.ensure_one()  # Asegurarse de que solo se procesa una encuesta a la vez

        # Obtener las respuestas completas
        completas = self.env['survey.user_input'].search([('survey_id', '=', self.id), ('state', '=', 'done')])
        if not completas:
            raise UserError('No hay respuestas completas para esta encuesta.')

        # Obtener líneas de respuestas
        respuestas = self.env['survey.user_input.line'].search([('survey_id', '=', self.id)])

        # Mapear preguntas específicas con nombres de columnas
        questions_map = {
            10: 'Nombres',
            11: 'Apellidos',
            12: 'DNI',
            13: 'Localidad',
            14: 'Dirección',
            15: 'Fecha de Nacimiento',
            16: 'Teléfono',
            17: 'Correo Electrónico',
            18: 'Preferencia de Voto',
        }

        # Crear el archivo Excel en memoria
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Resultados')

        # Escribir encabezados
        headers = list(questions_map.values())
        headers.insert(0, 'ID Usuario')  # Agregar el identificador del usuario
        for col_num, header in enumerate(headers):
            worksheet.write(0, col_num, header)

        # Escribir datos
        row = 1
        for completa in completas:
            # Inicializar datos del usuario
            user_data = [completa.id]  # Agregar el ID del usuario

            for question_id, column_name in questions_map.items():
                # Buscar la respuesta de esa pregunta específica
                respuesta = respuestas.filtered(lambda r: r.user_input_id.id == completa.id and r.question_id.id == question_id)
                user_data.append(respuesta.value_char_box if respuesta else '')

            # Escribir la fila de datos
            for col_num, cell_value in enumerate(user_data):
                worksheet.write(row, col_num, cell_value)
            row += 1

        # Cerrar el workbook
        workbook.close()
        output.seek(0)

        # Crear archivo adjunto para descargar
        attachment = self.env['ir.attachment'].create({
            'name': f'Resultados_Encuesta_{self.id}.xlsx',
            'type': 'binary',
            'datas': base64.b64encode(output.read()),
            'res_model': 'survey.survey',
            'res_id': self.id,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        })
        output.close()

        # Devolver acción para descargar el archivo
        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content/{attachment.id}?download=true',
            'target': 'new',
        }