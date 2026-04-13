# Maintenance Analytics — Présentation complète
## Application de suivi et d'analytique de maintenance industrielle | OCP Group — Site Daoui

---

## Slide 1 — Titre

# Maintenance Analytics
### Une application web interne de pilotage de la maintenance industrielle
**OCP Group — Site Daoui | Bureau de méthode**

> Développée pour centraliser, visualiser et piloter toutes les activités de maintenance du site en temps réel, directement connectée à SAP PM.

---

## Slide 2 — Contexte & Problématique

# Pourquoi cette application ?

### Avant Maintenance Analytics, le service méthode faisait face à :

- **Données SAP dispersées** : les données des ordres de travail, avis et pièces de rechange étaient accessibles uniquement via SAP, nécessitant des exports manuels et des mises à jour périodiques
- **Pas de visibilité en temps réel** : impossible de connaître instantanément le nombre d'OTs LANC, CRPR ou en retard sans générer un rapport SAP
- **Suivi PDR manuel** : la disponibilité des pièces de rechange était gérée par emails et appels téléphoniques, sans traçabilité
- **Demandes moteurs non structurées** : les demandes de moteurs électriques arrivaient de façon informelle (papier, WhatsApp, email)
- **Reporting chronophage** : la préparation des rapports hebdomadaires prenait plusieurs heures chaque semaine
- **Coordination difficile** entre les services : mécanique, électrique, instrumentation, installation, bureau de méthode

---

## Slide 3 — Objectifs de l'application

# Ce que Maintenance Analytics résout

### 4 objectifs principaux :

**1. Centralisation des données**
Connecter toutes les sources (SAP PM, Google Sheets) en un seul point d'accès visuel et interactif

**2. Pilotage en temps réel**
Tableaux de bord avec KPIs actualisables à la demande : OTs LANC, CRPR, en retard, par poste, par installation

**3. Digitalisation des processus**
- Confirmation PDR en ligne avec traçabilité
- Demandes de moteurs électriques structurées et suivies
- Validation et approbation avec workflow clair

**4. Communication automatisée**
- Rappels automatiques par email chaque mercredi
- Notifications pour les nouvelles demandes
- Chatbot IA pour répondre aux questions de pilotage instantanément

---

## Slide 4 — Architecture technique

# Comment ça marche ?

### Architecture 3 couches

**Couche données (Google Sheets)**
- Feuille "Travaux hebdomadaire" : OTs SAP exportés (statuts, postes, PDR, réalisation)
- Feuille "Users" : gestion des utilisateurs et profils
- Feuilles "Demandes", "Demande des interchangeables" : moteurs électriques

**Couche logique (Google Apps Script + Vercel)**
- Google Apps Script : lecture/écriture Sheets, envoi d'emails automatiques
- Vercel (serverless) : proxy sécurisé entre l'application et les Google Sheets

**Couche présentation (Application Web)**
- Application web monopage (HTML/CSS/JavaScript)
- Interface responsive, accessible depuis n'importe quel navigateur
- Aucune installation requise
- Chatbot IA intégré (Groq API)

---

## Slide 5 — Profils utilisateurs

# Qui utilise l'application ?

### 7 profils configurés avec des vues adaptées

| Profil | Accès | Responsabilité |
|--------|-------|----------------|
| **Admin** | Complet | Validation demandes, paramètres |
| **Responsable méthode** | Lecture complète + CC emails | Supervision globale |
| **Bureau de méthode** | PDR keywords spéciaux | Confirmation réducteurs, pompes, moteurs |
| **Appro mécanique** | PDR 421-MEC | Confirmation pièces mécaniques |
| **Appro électrique** | PDR 423-ELEC | Confirmation pièces électriques |
| **Appro installation** | PDR 421-INST | Confirmation pièces installation |
| **Appro Instrumentation** | PDR 423-REG | Confirmation pièces instrumentation |
| **Interchangeable électrique** | CC rappels, consultation | Suivi interchangeables |

> Chaque utilisateur voit uniquement les données qui le concernent — l'interface s'adapte automatiquement à son profil.

---

## Slide 6 — Tableau de bord (Dashboard)

# Vue d'ensemble en un coup d'œil

### KPIs affichés en temps réel

**Indicateurs principaux :**
- Nombre total d'OTs de la semaine
- OTs **LANC** (lancés, en cours d'exécution)
- OTs **CRPR** (en cours de préparation)
- OTs en **retard** (non clôturés hors délai)
- OTs **CONF** (confirmés / clôturés)
- Nombre d'**Avis** en attente (AOUV + AENC)

**Répartition visuelle :**
- Graphiques par poste de travail (421-MEC, 423-ELEC, 421-INST, 423-REG...)
- Répartition par installation
- Évolution hebdomadaire

**Bouton d'actualisation** en haut à droite : rafraîchit toutes les données sans quitter la page en cours

---

## Slide 7 — Module Rapport OTs SAP

# Suivi des Ordres de Travail

### Consultation complète des OTs de la semaine

**Filtres disponibles :**
- Par poste de travail (421-MEC, 423-ELEC, 421-INST, 423-REG, etc.)
- Par statut système (LANC, CRPR, CONF, CLOT...)
- Par statut utilisateur (ATPL, CRPR...)
- Par installation
- Recherche libre

**Informations affichées par OT :**
- N° d'ordre, objet technique, description
- Poste de travail, installation
- Statut SAP (système + utilisateur)
- PDR associée et statut de confirmation
- Réalisation (Fait / Non fait)

**Export et impression** disponibles pour les rapports hebdomadaires

---

## Slide 8 — Module Avis SAP

# Gestion des Avis de maintenance

### Suivi des avis en attente de traitement

**Statuts suivis :**
- **AOUV** : Avis ouvert, en attente de création d'OT
- **AENC** : Avis en cours de traitement
- Avis clôturés (historique)

**Indicateurs clés :**
- Nombre d'avis AOUV + AENC (en attente OT)
- Répartition par poste de travail
- Répartition par installation
- Liste des avis les plus anciens (priorité de traitement)

**Bénéfice :** Plus aucun avis "oublié" dans SAP — le service méthode visualise en temps réel tous les avis en attente de traitement sur l'ensemble du site.

---

## Slide 9 — Module Planification

# Planning de maintenance

### Vue planning de la semaine par poste de travail

**Contenu du planning :**
- OTs planifiés pour la semaine en cours
- Visualisation par jour et par poste
- Charge de travail par équipe

**Fonctionnalités :**
- Filtrage par poste, par installation, par statut
- Vue condensée / vue détaillée
- Indicateurs de charge : comparaison charge planifiée vs capacité

**Onglets de la page planification :**
- Vue Planning
- Suivi des réalisations
- **PDR (Pièces de Rechange)** — module dédié
- Arrêts planifiés

---

## Slide 10 — Module PDR (Pièces de Rechange)

# Le cœur opérationnel de l'application

### Problématique résolue
Avant : la disponibilité des PDR se gérait par téléphone et email, sans traçabilité. Les OTs étaient bloqués faute de confirmation à temps.

### Comment ça fonctionne
**1. Identification automatique** : chaque OT ayant une PDR renseignée dans la feuille SAP est affiché dans l'onglet PDR

**2. Affectation intelligente par poste :**
- 421-MEC → Service Appro mécanique
- 423-ELEC → Service Appro électrique
- 421-INST → Service Appro installation
- 423-REG → Service Appro Instrumentation
- PDR contenant *réducteur, pompe, moteur* → **Bureau de méthode**

**3. Statuts de confirmation :**
- ✅ **En stock** : pièce disponible
- ❌ **Non disponible** : pièce manquante (avec délai estimé)
- 💬 **Observation** : information partielle, complément nécessaire
- ⏳ **En attente** : pas encore de réponse

---

## Slide 11 — PDR — Workflow de confirmation

# De la demande à la confirmation

### Étapes du processus

```
SAP PM → Export hebdomadaire → Google Sheets
                                    ↓
                          Maintenance Analytics détecte les OTs avec PDR
                                    ↓
                    Affichage dans l'onglet PDR du service concerné
                                    ↓
              Service Appro ouvre la carte → confirme la disponibilité
                          ↓              ↓              ↓
                    En stock       Non disponible    Observation
                    (col T = OUI)  (col T = NON)   (col U remplie)
                          ↓
              Donnée sauvegardée dans Google Sheets (colonne T/U)
                          ↓
              Visible immédiatement par le bureau de méthode
```

**Nouveauté : champ Observation**
Le service appro peut saisir une observation (besoin d'info complémentaire) **sans être obligé de choisir un statut** — la coordination continue sans blocage.

---

## Slide 12 — PDR — Rappels automatiques par email

# Ne plus oublier de confirmer les PDR

### Système d'emails automatiques

**Déclenchement :** Chaque **mercredi à 08h00** (Google Apps Script trigger)

**Logique d'envoi :**
- Le script scanne tous les OTs actifs (LANC/CRPR) ayant une PDR sans confirmation
- Envoie un email récapitulatif **uniquement si des PDR sont en attente**
- Pas d'email envoyé si tout est confirmé

**Destinataires par poste :**
| Poste | Service destinataire |
|-------|---------------------|
| 421-MEC | Appro mécanique |
| 423-ELEC | Appro électrique |
| 421-INST | Appro installation |
| 423-REG | Appro Instrumentation |

**CC systématique :** Responsable méthode + Interchangeable électrique sur tous les emails

**Contenu de l'email :** tableau récapitulatif avec N° OT, objet technique, description, PDR demandée

---

## Slide 13 — Module Demandes Moteurs Électriques

# Digitalisation des demandes de remplacement

### Avant
Les demandes de moteurs arrivaient de façon informelle : papier, appels téléphoniques, messages. Pas de suivi, pas d'historique.

### Maintenant — Formulaire structuré
**Informations saisies par le demandeur :**
- Type de demande : **Pose** (nouveau) ou **Réparation**
- Installation et objet technique
- Caractéristiques du moteur : puissance (kW), tension (V), vitesse (tr/min)
- Description de l'anomalie
- Matricule et nom du demandeur

**Workflow de validation :**
- **Pose** : nécessite approbation Admin → email de notification envoyé
- **Réparation** : approbation automatique (urgence opérationnelle)

**Suivi en temps réel :**
- Statut : En attente / Approuvé / Refusé
- Historique complet avec dates
- Justification en cas de refus

---

## Slide 14 — Module Arrêts Planifiés

# Suivi des arrêts de production

### Fonctionnalités

**Calendrier des arrêts :**
- Vue mensuelle des arrêts planifiés par installation
- Durée, type d'arrêt, équipement concerné
- Statut : planifié / en cours / terminé / reporté

**Marquage "Reporté" :**
- Un arrêt peut être marqué comme reporté directement depuis l'application
- La donnée est mise à jour dans le fichier Planning Google Sheets
- Traçabilité complète avec date de report

**Indicateurs :**
- Nombre d'arrêts planifiés sur la période
- Taux de réalisation vs planification
- Arrêts reportés (suivi des glissements)

---

## Slide 15 — Chatbot IA intégré

# Un assistant intelligent disponible 24h/24

### Powered by Groq (LLaMA 3)

**Ce que le chatbot connaît en temps réel :**
- Tous les KPIs du tableau de bord
- Répartition des OTs LANC et CRPR par poste de travail
- Statut des PDR (en attente, confirmées, observations)
- Données des avis SAP (AOUV, AENC)
- Planning de la semaine

**Exemples de questions posées :**
- *"Combien d'OTs sont en attente de confirmation PDR ?"*
- *"Donne-moi la répartition des OTs LANC par corps de métier"*
- *"Quels sont les OTs CRPR du poste 421-MEC ?"*
- *"Combien d'avis sont en attente de traitement ?"*

**Interface :**
- Bouton flottant (icône robot animé) accessible depuis toutes les pages
- Réponse en langage naturel en français
- Contexte mis à jour à chaque actualisation des données

---

## Slide 16 — Notifications Push

# Rester informé en temps réel

### Système de notifications web

**Déclencheurs de notification :**
- Nouvelle demande de moteur électrique soumise → Admin notifié
- Demande approuvée ou refusée → Demandeur notifié
- PDR confirmée → Notification bureau de méthode

**Fonctionnement :**
- Notifications push navigateur (même si l'onglet est fermé)
- Pas d'application mobile nécessaire
- Fonctionne sur PC et mobile
- Abonnement par profil utilisateur

**Avantage :** Le service méthode est informé instantanément de toute nouvelle demande ou action, sans avoir à consulter l'application en permanence.

---

## Slide 17 — Suivi des Réalisations

# Clôture des OTs en temps réel

### Marquage Fait / Non Fait

**Fonctionnalité :**
Chaque OT planifié peut être marqué **Fait** ou **Non Fait** directement depuis l'application, sans passer par SAP.

**Processus :**
- Le technicien ou chef d'équipe ouvre l'OT dans l'application
- Sélectionne : ✅ Fait / ❌ Non Fait
- La valeur est enregistrée dans Google Sheets (colonne O)
- Visible immédiatement par le bureau de méthode

**Bénéfice :**
- Taux de réalisation hebdomadaire calculé automatiquement
- Identification rapide des OTs non réalisés
- Rapport de réalisation généré sans saisie manuelle supplémentaire

---

## Slide 18 — Sécurité et accès

# Une application sécurisée

### Mesures en place

**Authentification :**
- Accès par profil utilisateur géré dans la feuille Users
- Chaque utilisateur voit uniquement les données de son périmètre

**Sécurité des données :**
- Feuille Users : accès restreint (non public)
- Code source : dépôt GitHub privé
- Clés API gérées côté serveur (non exposées dans le code)
- Proxy Vercel : aucune donnée sensible exposée côté client

**Infrastructure :**
- Hébergement Vercel (HTTPS obligatoire)
- Google Apps Script : exécution avec compte autorisé uniquement
- Aucune donnée stockée en dehors de l'environnement OCP Google Workspace

---

## Slide 19 — Bénéfices & Gains mesurables

# Ce que ça change concrètement

### Gains de temps

| Tâche | Avant | Après |
|-------|-------|-------|
| Rapport hebdomadaire OTs | 2-3 heures | 5 minutes (actualisation) |
| Suivi PDR par poste | 30 min d'emails/appels | Temps réel sur l'écran |
| Demande de moteur | Formulaire papier + saisie manuelle | Formulaire numérique direct |
| Rappel confirmation PDR | Manuel chaque semaine | Automatique chaque mercredi |
| Réponse aux questions de pilotage | Consultation SAP | Chatbot instantané |

### Gains qualitatifs
- **Traçabilité complète** de toutes les actions (PDR, demandes, réalisations)
- **Coordination améliorée** entre les services appro, méthode et terrain
- **Réactivité augmentée** grâce aux notifications et rappels automatiques
- **Décision facilitée** avec les KPIs visuels en temps réel
- **Zéro email informel** pour les demandes et confirmations PDR

---

## Slide 20 — Conclusion & Perspectives

# Une base solide, des évolutions possibles

### Ce qui est opérationnel aujourd'hui
- ✅ Tableau de bord temps réel (OTs, Avis, KPIs)
- ✅ Suivi PDR avec confirmation multi-service
- ✅ Rappels automatiques par email (4 services + CC)
- ✅ Demandes moteurs électriques avec workflow de validation
- ✅ Chatbot IA de pilotage
- ✅ Notifications push
- ✅ Suivi des réalisations
- ✅ Calendrier des arrêts planifiés

### Évolutions envisageables
- 📊 Tableaux de bord analytiques avancés (tendances, historique multi-semaines)
- 📱 Application mobile native
- 🔗 Connexion directe SAP (sans export manuel)
- 📄 Génération automatique de rapports PDF hebdomadaires
- 🔔 Escalade automatique pour PDR non confirmées après X jours
- 📦 Module gestion de stock pièces de rechange

---

*Maintenance Analytics — Développé par le Bureau de méthode Daoui | OCP Group*
*Questions ? Contactez : m.elamraoui@ocpgroup.ma*
