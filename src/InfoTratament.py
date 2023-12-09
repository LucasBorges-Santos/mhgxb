import pandas as pd
import json
from docxtpl import DocxTemplate
import aspose.words as aw


class InfoTratament:
    context = {}
    def __init__(self, data_path:str, template_path:str) -> None:
        self.data_path = data_path
        self.template_path = template_path
        self.get_data()
        self.get_template()
        
    def get_data(self):
        self.df_data = pd.read_excel(self.data_path)
    
    def get_template(self):
        self.doc = DocxTemplate(self.template_path)
        
    def render_data(self, context):
        self.doc.render(context)
        
    def save_data(self, save_path:str):
        self.doc.save(save_path)
    
    def render_save(self, context:dict, save_path:str):
        self.get_template()
        self.render_data(context)
        self.save_data(save_path)
    
    def get_template_data(self, current_data:dict) -> dict:
        context = {}
        context['id'] =                   current_data["ID"]
        context['nota_inspecao'] =        current_data["11. Inspeção de Equipamentos( 1 - 10 )"]
        context['equipamentos'] =         json.loads(current_data["8. Preparação para Emergências (todas as opções que se aplicam)"])
        context['nota_avaliacao_risco'] = json.loads(current_data["14. Avaliação de Riscos (todas as opções que se aplicam)"])
        context['nota_limpeza'] =         current_data["3.  Limpeza, organização e manutenção ( 1 - 10 )"]
        context['nota_proc_seg'] =        current_data["5. Procedimentos de Segurança: ( 1 - 10 )"]
        context['epis'] =                 json.loads(current_data["2.  EPIs utilizados  (todas as opções que se aplicam):"])
        context['sinalizacao'] =          json.loads(current_data["4. Sinalização de Segurança (todas as opções que se aplicam)"])
        context['nota_iluminacao'] =      current_data["9. Iluminação e Ventilação: ( 1 - 10 )"]
        context['ergon'] =                json.loads(current_data["10. Ergonomia( todas as opções que se aplicam)"])
        context['treinamentos'] =         json.loads(current_data["6. Treinamento em segurança(todas as opções que se aplicam)"])
        context['nota_part'] =            current_data["13. Participação dos Trabalhadores ( 1 - 10 )"]
        context['hist'] =                 json.loads(current_data["12. Registro de Incidentes(todas as opções que se aplicam)"])
        context['nota_geral'] =           current_data["1.  Segurança geral ( 1 - 10 )"]
        return context
        
    def execute_data(self):
        self.df_data.columns = [x.replace("\n", "") for x in self.df_data.columns]
        data = self.df_data.T.to_dict()
        
        for id, current_data in data.items():
            current_data = self.get_template_data(current_data)
            self.render_save(current_data, f'../laudos/laudo_{current_data["id"]}.docx')
            doc = aw.Document(f'../laudos/laudo_{current_data["id"]}.docx')
            doc.save(f'../laudos/laudo_{current_data["id"]}.pdf')

if __name__ == "__main__":
    InfoTratament(r"../data/data.xlsx", r"../template/template.docx").execute_data()