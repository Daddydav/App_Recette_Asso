import os, json, xlsxwriter
from datetime import datetime
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.metrics import dp
from kivy.utils import platform
from kivy.uix.scrollview import ScrollView

class LotoUniversel(BoxLayout):
    def __init__(self, **kwargs):
        super(LotoUniversel, self).__init__(**kwargs)
        self.orientation = 'vertical'
        self.padding = [dp(15), dp(40), dp(15), dp(15)]
        self.spacing = dp(10)

        self.path_json, self.path_xlsx, self.path_txt = self.get_all_paths()
        
        # Chargement des produits
        self.produits = self.charger_config_json()
        
        # Initialisation des compteurs (Dynamique)
        self.cur_stats = {p['id']: 0 for p in self.produits}
        self.glob_stats = {p['id']: 0 for p in self.produits}
        self.grand_total_argent = 0.0
        self.cur_prix_vente = 0.0

        self.restaurer_session_txt()
        self.create_interface()

    def get_all_paths(self):
        date_s = datetime.now().strftime('%Y-%m-%d')
        if platform == 'android':
            from android.storage import primary_external_storage_path
            base = os.path.join(primary_external_storage_path(), 'Download')
        else:
            base = os.getcwd()
        return (os.path.join(base, 'produits.json'), 
                os.path.join(base, f"Ventes_{date_s}.xlsx"), 
                os.path.join(base, f".session_{date_s}.txt"))

    def charger_config_json(self):
        if os.path.exists(self.path_json):
            try:
                with open(self.path_json, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
                    if not content: # Si le fichier est vide
                        print("Le fichier JSON est vide !")
                        return self.default_config()
                    return json.loads(content)
            except json.JSONDecodeError:
                print("Erreur de formatage dans le JSON !")
                return self.default_config()
        return self.default_config()

    def default_config(self):
        return [{"id": "item1", "nom": "Produit de Secours", "prix": 1.0}]

    def restaurer_session_txt(self):
        if os.path.exists(self.path_txt):
            try:
                with open(self.path_txt, "r") as f:
                    data = json.load(f)
                    self.glob_stats = data.get('stats', self.glob_stats)
                    self.grand_total_argent = data.get('total', 0.0)
            except: pass

    def create_interface(self):
        self.clear_widgets()
        
        # Header global
        self.add_widget(Label(text=f"SESSION CUMULÉE : {self.grand_total_argent:.2f}€", 
                             size_hint_y=None, height=dp(30), font_size='16sp', color=(0.8,0.8,0.8,1)))

        # Zone de défilement pour les produits (si la liste est longue)
        scroll = ScrollView(size_hint=(1, 1))
        layout_produits = BoxLayout(orientation='vertical', size_hint_y=None, spacing=dp(5))
        layout_produits.bind(minimum_height=layout_produits.setter('height'))

        for p in self.produits:
            row = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(75), spacing=dp(10))
            
            # Label Nom + Total cumulé discret
            # Avant (plante si l'ID n'existe pas dans les stats)
            total_cumul = self.glob_stats.get(p['id'], 0)
            txt_prod = f"{p['nom']}\n({p['prix']}€)\nTotal: {total_cumul}"
            row.add_widget(Label(text=txt_prod, size_hint_x=0.4, font_size='18sp', bold=True, halign='center'))
            
            # Boutons +/-
            btn_m = Button(text="-", size_hint_x=0.2, font_size='25sp', bold=True, background_color=(1.0, 0.0, 0.0, 1))
            btn_m.bind(on_press=lambda x, pid=p['id']: self.modifier_quantite(pid, -1))
            
            lbl_qte = Label(text="0", size_hint_x=0.2, font_size='24sp', bold=True)
            setattr(self, f"lbl_cur_{p['id']}", lbl_qte)
            
            btn_p = Button(text="+", size_hint_x=0.2, font_size='25sp', bold=True, background_color=(0.0, 1.0, 0.0, 1))
            btn_p.bind(on_press=lambda x, pid=p['id']: self.modifier_quantite(pid, 1))
            
            row.add_widget(btn_m); row.add_widget(lbl_qte); row.add_widget(btn_p)
            layout_produits.add_widget(row)
        
        scroll.add_widget(layout_produits)
        self.add_widget(scroll)

        # Affichage Prix Vente en cours
        self.lbl_payer = Label(text="À PAYER : 0.00 €", font_size='38sp', bold=True,
                              size_hint_y=None, height=dp(90), color=(0.2, 0.6, 1, 1))
        self.add_widget(self.lbl_payer)

        # Bouton Valider
        btn_v = Button(text="VALIDER LA VENTE", size_hint_y=None, height=dp(80),
                      background_color=(0.1, 0.7, 0.2, 1), font_size='22sp', bold=True)
        btn_v.bind(on_press=self.valider_vente)
        self.add_widget(btn_v)

    def modifier_quantite(self, pid, delta):
        if self.cur_stats[pid] + delta >= 0:
            self.cur_stats[pid] += delta
            getattr(self, f"lbl_cur_{pid}").text = str(self.cur_stats[pid])
            # Recalcul du prix total de la vente
            self.cur_prix_vente = sum(self.cur_stats[p['id']] * p['prix'] for p in self.produits)
            self.lbl_payer.text = f"À PAYER : {self.cur_prix_vente:.2f} €"

    def valider_vente(self, instance):
        if self.cur_prix_vente == 0: return
        
        # Transfert vers Global
        for p in self.produits:
            self.glob_stats[p['id']] += self.cur_stats[p['id']]
            self.cur_stats[p['id']] = 0 # Reset
            
        self.grand_total_argent += self.cur_prix_vente
        self.cur_prix_vente = 0.0
        
        self.sauvegarde_fichiers()
        self.create_interface()

    def sauvegarde_fichiers(self):
        # 1. Sauvegarde TXT (Format JSON pour être plus robuste)
        try:
            with open(self.path_txt, "w") as f:
                json.dump({'stats': self.glob_stats, 'total': self.grand_total_argent}, f)
        except: pass

        # 2. Sauvegarde EXCEL
        try:
            wb = xlsxwriter.Workbook(self.path_xlsx)
            ws = wb.add_worksheet("Ventes")
            headers = ["PRODUIT", "QUANTITÉ", "PRIX UNIT.", "TOTAL"]
            for c, h in enumerate(headers): ws.write(0, c, h)
            
            row_idx = 1
            for p in self.produits:
                qte = self.glob_stats[p['id']]
                ws.write(row_idx, 0, p['nom'])
                ws.write(row_idx, 1, qte)
                ws.write(row_idx, 2, p['prix'])
                ws.write(row_idx, 3, qte * p['prix'])
                row_idx += 1
            
            ws.write(row_idx + 1, 2, "TOTAL GÉNÉRAL")
            ws.write(row_idx + 1, 3, self.grand_total_argent)
            wb.close()
        except: pass

class LotoApp(App):
    def build(self):
        self.icon = "icon.png"
        return LotoUniversel()

if __name__ == '__main__':
    LotoApp().run()