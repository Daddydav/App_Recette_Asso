import os, xlsxwriter, json, math
from datetime import datetime
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.gridlayout import GridLayout
from kivy.metrics import dp
from kivy.utils import platform
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup
import configparser

if platform == 'android':
    from android.permissions import request_permissions, Permission
    request_permissions([Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE])
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
        self.historique_ventes = []

        self.restaurer_session_txt()
        self.create_interface()

    def charger_fich_conf(self):
        # Chemin vers le fichier ini dans Download
        if platform == 'android':
            from android.storage import primary_external_storage_path
            path_ini = os.path.join(primary_external_storage_path(), 'Download', 'configuration.ini')
        else:
            path_ini = 'configuration.ini'

        config = configparser.ConfigParser()

        if os.path.exists(path_ini):
            config.read(path_ini)
            # On récupère les valeurs (en s'assurant qu'elles sont converties en float)
            try:
                s = config.get('Config', 'fich_conf', fallback='produits.json')
                print("Succès", f"Fichier chargé depuis :\n{path_ini}")
            except Exception as e:
                print(f"Erreur de lecture ini, utilisation des prix par défaut: {e}")
                s = 'produits.json'
        else:
            s = 'produits.json'

        return s

    def get_all_paths(self):
        date_s = datetime.now().strftime('%Y-%m-%d')
        if platform == 'android':
            from android.storage import primary_external_storage_path
            base = os.path.join(primary_external_storage_path(), 'Download')
        else:
            base = os.getcwd()

        return (os.path.join(base, self.charger_fich_conf()), 
                os.path.join(base, f"Ventes_{date_s}.xlsx"), 
                os.path.join(base, f".session_{date_s}.txt"))
    
    def charger_config_json(self):
        if os.path.exists(self.path_json):
            try:
                with open(self.path_json, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
                    if not content:
                        return self.default_config()
                    return json.loads(content)
            except PermissionError:
                print("ERREUR : Accès refusé par Android. Vérifiez les permissions !")
                return self.default_config()
            except Exception as e:
                print(f"Erreur lors de la lecture : {e}")
                return self.default_config()
        return self.default_config()

    def default_config(self):
        return [{"id": "item1", "nom": "Produit de Secours", "prix": 1.0}]

    def restaurer_session_txt(self):
        if os.path.exists(self.path_txt):
            try:
                with open(self.path_txt, "r") as f:
                    data = json.load(f)
                    self.grand_total_argent = data.get('total', 0.0)
                    historique = data.get('historique_ventes')
                    if historique is not None:
                        self.historique_ventes = historique
                        self.glob_stats = {}
                        for ligne in historique:
                            pid = ligne['id']
                            self.glob_stats[pid] = self.glob_stats.get(pid, 0) + ligne['quantite']
                    else:
                        self.glob_stats.update(data.get('stats', {}))
                        produits_par_id = {p['id']: p for p in self.produits}
                        for pid, quantite in self.glob_stats.items():
                            if quantite > 0:
                                produit = produits_par_id.get(pid, {
                                    'nom': f"Produit retiré ({pid})",
                                    'prix': 0.0
                                })
                                self.historique_ventes.append({
                                    'id': pid,
                                    'nom': produit['nom'],
                                    'prix': produit['prix'],
                                    'quantite': quantite
                                })
            except (OSError, ValueError, TypeError, KeyError) as e:
                print(f"Erreur lors de la restauration de la session : {e}")
        for produit in self.produits:
            self.glob_stats.setdefault(produit['id'], 0)

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
            row = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(65), spacing=dp(10))
            
            # Label Nom + Total cumulé discret
            # Avant (plante si l'ID n'existe pas dans les stats)
            total_cumul = self.glob_stats.get(p['id'], 0)
            txt_prod = f"{p['nom']}\n({p['prix']}€)\nTotal: {total_cumul}"
            row.add_widget(Label(text=txt_prod, size_hint_x=0.4, font_size='18sp', bold=True, halign='center'))
            
            # Boutons +/-
            btn_m = Button(text="-", size_hint_x=0.2, font_size='25sp', bold=True, background_color=(1.0, 0.0, 0.0, 1))
            btn_m.bind(on_press=lambda x, pid=p['id']: self.modifier_quantite(pid, -1))
            
            lbl_qte = Label(text=str(self.cur_stats.get(p['id'], 0)),
                            size_hint_x=0.2, font_size='24sp', bold=True)
            setattr(self, f"lbl_cur_{p['id']}", lbl_qte)
            
            btn_p = Button(text="+", size_hint_x=0.2, font_size='25sp', bold=True, background_color=(0.0, 1.0, 0.0, 1))
            btn_p.bind(on_press=lambda x, pid=p['id']: self.modifier_quantite(pid, 1))
            
            row.add_widget(btn_m); row.add_widget(lbl_qte); row.add_widget(btn_p)
            layout_produits.add_widget(row)
        
        scroll.add_widget(layout_produits)
        self.add_widget(scroll)

        # Affichage Prix Vente en cours
        self.lbl_payer = Label(text=f"À PAYER : {self.cur_prix_vente:.2f} €",
                              font_size='38sp', bold=True,
                              size_hint_y=None, height=dp(90), color=(0.2, 0.6, 1, 1))
        self.add_widget(self.lbl_payer)

        # Bouton Valider
        btn_v = Button(text="VALIDER LA VENTE", size_hint_y=None, height=dp(80),
                      background_color=(0.1, 0.7, 0.2, 1), font_size='22sp', bold=True)
        btn_v.bind(on_press=self.valider_vente)
        self.add_widget(btn_v)

        btn_produits = Button(text="GÉRER LES PRODUITS", size_hint_y=None, height=dp(55),
                              background_color=(0.2, 0.35, 0.65, 1), font_size='18sp', bold=True)
        btn_produits.bind(on_press=self.ouvrir_editeur_produits)
        self.add_widget(btn_produits)

    def ouvrir_editeur_produits(self, instance):
        contenu = BoxLayout(orientation='vertical', spacing=dp(8), padding=dp(10))
        scroll = ScrollView()
        lignes = GridLayout(cols=1, spacing=dp(5), size_hint_y=None)
        lignes.bind(minimum_height=lignes.setter('height'))
        champs = []

        def ajouter_ligne(produit):
            ligne = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(48), spacing=dp(5))
            nom = TextInput(text=produit['nom'], hint_text="Nom", multiline=False,
                            size_hint_x=0.58)
            prix = TextInput(text=str(produit['prix']), hint_text="Prix", multiline=False,
                             input_filter='float', size_hint_x=0.25)
            bouton_supprimer = Button(text="X", size_hint_x=0.17,
                                      background_color=(0.8, 0.2, 0.2, 1))
            entree = {'id': produit['id'], 'nom': nom, 'prix': prix, 'ligne': ligne}
            bouton_supprimer.bind(on_press=lambda _btn, e=entree: supprimer_ligne(e))
            ligne.add_widget(nom)
            ligne.add_widget(prix)
            ligne.add_widget(bouton_supprimer)
            lignes.add_widget(ligne)
            champs.append(entree)

        def supprimer_ligne(entree):
            lignes.remove_widget(entree['ligne'])
            champs.remove(entree)

        for produit in self.produits:
            ajouter_ligne(produit)

        scroll.add_widget(lignes)
        contenu.add_widget(scroll)

        btn_ajouter = Button(text="AJOUTER UN PRODUIT", size_hint_y=None, height=dp(45))
        contenu.add_widget(btn_ajouter)

        actions = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(50), spacing=dp(8))
        btn_annuler = Button(text="ANNULER")
        btn_enregistrer = Button(text="ENREGISTRER", background_color=(0.1, 0.7, 0.2, 1))
        actions.add_widget(btn_annuler)
        actions.add_widget(btn_enregistrer)
        contenu.add_widget(actions)

        popup = Popup(title="Modifier les produits", content=contenu,
                      size_hint=(0.95, 0.9), auto_dismiss=False)

        def ajouter_nouveau_produit(_btn):
            ids = ({entree['id'] for entree in champs}
                   | set(self.glob_stats)
                   | {ligne['id'] for ligne in self.historique_ventes})
            numero = 1
            while f"item{numero}" in ids:
                numero += 1
            ajouter_ligne({'id': f"item{numero}", 'nom': '', 'prix': 0.0})

        def enregistrer(_btn):
            nouveaux_produits = []
            try:
                for entree in champs:
                    nom = entree['nom'].text.strip()
                    prix_texte = entree['prix'].text.strip().replace(',', '.')
                    if not nom:
                        raise ValueError("Chaque produit doit avoir un nom.")
                    prix = float(prix_texte)
                    if not math.isfinite(prix) or prix < 0:
                        raise ValueError("Les prix doivent être des nombres positifs ou nuls.")
                    nouveaux_produits.append({
                        'id': entree['id'],
                        'nom': nom,
                        'prix': prix
                    })

                if not nouveaux_produits:
                    raise ValueError("La liste doit contenir au moins un produit.")

                nouveaux_ids = {p['id'] for p in nouveaux_produits}
                produits_supprimes_avec_vente = [
                    p['nom'] for p in self.produits
                    if p['id'] not in nouveaux_ids and self.cur_stats.get(p['id'], 0) > 0
                ]
                if produits_supprimes_avec_vente:
                    raise ValueError(
                        "Terminez ou annulez les quantités en cours avant de supprimer : "
                        + ", ".join(produits_supprimes_avec_vente)
                    )

                self.sauvegarder_session_txt()
                self.sauvegarder_config_json(nouveaux_produits)
            except (OSError, ValueError, TypeError) as e:
                self.afficher_message("Enregistrement impossible", str(e))
                return

            anciennes_quantites = self.cur_stats
            self.produits = nouveaux_produits
            self.cur_stats = {
                p['id']: anciennes_quantites.get(p['id'], 0) for p in self.produits
            }
            for p in self.produits:
                self.glob_stats.setdefault(p['id'], 0)
            self.cur_prix_vente = sum(
                self.cur_stats[p['id']] * p['prix'] for p in self.produits
            )
            popup.dismiss()
            self.create_interface()

        btn_ajouter.bind(on_press=ajouter_nouveau_produit)
        btn_annuler.bind(on_press=lambda _btn: popup.dismiss())
        btn_enregistrer.bind(on_press=enregistrer)
        popup.open()

    def sauvegarder_config_json(self, produits):
        chemin_temporaire = self.path_json + ".tmp"
        try:
            with open(chemin_temporaire, 'w', encoding='utf-8') as f:
                json.dump(produits, f, ensure_ascii=False, indent=4)
            os.replace(chemin_temporaire, self.path_json)
        except Exception:
            if os.path.exists(chemin_temporaire):
                try:
                    os.remove(chemin_temporaire)
                except OSError as e:
                    print(f"Erreur de nettoyage du fichier temporaire : {e}")
            raise

    def sauvegarder_session_txt(self):
        chemin_temporaire = self.path_txt + ".tmp"
        try:
            with open(chemin_temporaire, "w", encoding='utf-8') as f:
                json.dump({
                    'stats': self.glob_stats,
                    'total': self.grand_total_argent,
                    'historique_ventes': self.historique_ventes
                }, f, ensure_ascii=False)
            os.replace(chemin_temporaire, self.path_txt)
        except Exception:
            if os.path.exists(chemin_temporaire):
                try:
                    os.remove(chemin_temporaire)
                except OSError as e:
                    print(f"Erreur de nettoyage du fichier temporaire : {e}")
            raise

    def afficher_message(self, titre, message):
        contenu = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        contenu.add_widget(Label(text=message))
        bouton = Button(text="OK", size_hint_y=None, height=dp(45))
        contenu.add_widget(bouton)
        popup = Popup(title=titre, content=contenu, size_hint=(0.9, 0.4))
        bouton.bind(on_press=popup.dismiss)
        popup.open()

    def modifier_quantite(self, pid, delta):
        if self.cur_stats[pid] + delta + self.glob_stats.get(pid, 0) >= 0: # On ne peut pas vendre plus que ce qu'on a
            self.cur_stats[pid] += delta
            getattr(self, f"lbl_cur_{pid}").text = str(self.cur_stats[pid])
            # Recalcul du prix total de la vente
            self.cur_prix_vente = sum(self.cur_stats[p['id']] * p['prix'] for p in self.produits)
            self.lbl_payer.text = f"À PAYER : {self.cur_prix_vente:.2f} €"

    def valider_vente(self, instance):
        # On vérifie si on a au moins un produit à vendre
        #if self.cur_prix_vente <= 0:
        #    return

        # On transfère les compteurs temporaires vers les compteurs globaux
        for p in self.produits:
            p_id = p['id']
            # C'est ici que l'erreur 'c10' arrivait si on forçait l'ID
            quantite = self.cur_stats[p_id]
            self.glob_stats[p_id] = self.glob_stats.get(p_id, 0) + quantite
            if quantite > 0:
                ligne_existante = next((
                    ligne for ligne in self.historique_ventes
                    if ligne['id'] == p_id and ligne['nom'] == p['nom']
                    and ligne['prix'] == p['prix']
                ), None)
                if ligne_existante:
                    ligne_existante['quantite'] += quantite
                else:
                    self.historique_ventes.append({
                        'id': p_id,
                        'nom': p['nom'],
                        'prix': p['prix'],
                        'quantite': quantite
                    })
            self.cur_stats[p_id] = 0 # On remet à zéro pour la prochaine vente
            
        self.grand_total_argent += self.cur_prix_vente
        self.cur_prix_vente = 0.0
        
        # On sauvegarde les changements
        self.sauvegarde_fichiers()
        
        # On rafraîchit l'affichage
        self.create_interface()

    def sauvegarde_fichiers(self):
        # 1. Sauvegarde TXT (Format JSON pour être plus robuste)
        try:
            self.sauvegarder_session_txt()
        except OSError as e:
            print(f"Erreur lors de la sauvegarde de la session : {e}")

        # 2. Sauvegarde EXCEL
        try:
            wb = xlsxwriter.Workbook(self.path_xlsx)
            fmt_money = wb.add_format({'num_format': '#,##0.00 €'})
            fmt_bold = wb.add_format({'bold': True})
            fmt_res_money = wb.add_format({'bold': True, 'bg_color': '#D3D3D3','num_format': '#,##0.00 €'})
            fmt_res = wb.add_format({'bold': True, 'bg_color': '#D3D3D3'})
            ws = wb.add_worksheet("Ventes")
            headers = ["PRODUIT", "QUANTITÉ", "PRIX UNIT.", "TOTAL"]
            for c, h in enumerate(headers): ws.write(0, c, h, fmt_bold)
            
            row_idx = 1
            for ligne in self.historique_ventes:
                ws.write(row_idx, 0, ligne['nom'])
                ws.write(row_idx, 1, ligne['quantite'])
                ws.write(row_idx, 2, ligne['prix'], fmt_money)
                ws.write(row_idx, 3, ligne['quantite'] * ligne['prix'], fmt_money)
                row_idx += 1
            
            ws.write(row_idx + 1, 2, "TOTAL GÉNÉRAL", fmt_res)
            ws.write(row_idx + 1, 3, self.grand_total_argent, fmt_res_money)
            wb.close()
        except Exception as e:
            print(f"Erreur lors de la sauvegarde du fichier Excel : {e}")
class LotoApp(App):
    def build(self):
        if platform == 'android':
            request_permissions([
                Permission.READ_EXTERNAL_STORAGE, 
                Permission.WRITE_EXTERNAL_STORAGE,
                Permission.MANAGE_EXTERNAL_STORAGE
            ])
        self.icon = "icon.png"
        return LotoUniversel()
    
if __name__ == '__main__':
    LotoApp().run()