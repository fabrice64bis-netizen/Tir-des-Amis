# -*- coding: utf-8 -*-
"""
Utilitaires pour l'export en PDF avec support complet
"""

from pathlib import Path
from typing import List, Dict, Optional
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from datetime import datetime
from database.models import Database, Shooter
from config.settings import OUTPUT_DIR, A4_FORMAT
from utils.calculations import calculate_ranking, group_by_category, group_by_society


class PDFExporter:
    """Classe pour générer les PDFs"""
    
    def __init__(self, db: Database):
        self.db = db
        self.output_dir = OUTPUT_DIR
        self.output_dir.mkdir(exist_ok=True)
    
    def generate_general_ranking(self, filename: str = None) -> Path:
        """
        Génère le classement général (A4 Portrait)
        Highlight: ROI DU TIR
        """
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"01_Classement_General_{timestamp}.pdf"
        
        output_path = self.output_dir / filename
        
        # Récupérer tous les tireurs
        shooters = self.db.get_all_shooters()
        
        # Trier par score décroissant
        shooters_sorted = sorted(shooters, key=lambda x: x.points if hasattr(x, 'points') else 0, reverse=True)
        
        # Créer le document
        doc = SimpleDocTemplate(
            str(output_path),
            pagesize=A4,
            topMargin=10*mm,
            bottomMargin=10*mm,
            leftMargin=10*mm,
            rightMargin=10*mm,
            title="Classement Général"
        )
        
        elements = []
        styles = getSampleStyleSheet()
        
        # Titre
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=24,
            textColor=colors.HexColor('#003366'),
            spaceAfter=15,
            alignment=1,
            fontName='Helvetica-Bold'
        )
        
        elements.append(Paragraph("🏆 CLASSEMENT GÉNÉRAL 🏆", title_style))
        elements.append(Spacer(1, 5*mm))
        
        # Date
        date_text = datetime.now().strftime("%d/%m/%Y à %H:%M")
        date_style = ParagraphStyle(
            'DateStyle',
            parent=styles['Normal'],
            fontSize=10,
            alignment=2,
            textColor=colors.grey
        )
        elements.append(Paragraph(f"Généré le {date_text}", date_style))
        elements.append(Spacer(1, 10*mm))
        
        # Préparer les données du tableau
        data = [['Rang', 'Nom', 'Prénom', 'Catégorie', 'Société', 'Score', 'Série', '10s']]
        
        for rank, shooter in enumerate(shooters_sorted, 1):
            # Déterminer la couleur de fond
            is_roi = rank == 1  # Le premier est le ROI DU TIR
            
            score = getattr(shooter, 'score', 0) or 0
            serie = getattr(shooter, 'serie', 0) or 0
            n10 = getattr(shooter, 'n10', 0) or 0
            
            row = [
                str(rank),
                shooter.nom,
                shooter.prenom,
                shooter.categorie,
                shooter.societe or '',
                f"{score:.1f}" if isinstance(score, (int, float)) else str(score),
                str(serie),
                str(n10)
            ]
            data.append(row)
        
        # Créer le tableau
        table = Table(data, colWidths=[12*mm, 30*mm, 30*mm, 25*mm, 35*mm, 18*mm, 18*mm, 12*mm])
        
        # Style du tableau
        style = TableStyle([
            # En-tête
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#003366')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            
            # Corps du tableau
            ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),\n            ('FONTSIZE', (0, 1), (-1, -1), 9),\n            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),\n            \n            # Alternance de couleurs\n            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F5F5F5')]),\n            \n            # Premier rang (ROI DU TIR)\n            ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#FFD700')),\n            ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),\n            ('TEXTCOLOR', (0, 1), (-1, 1), colors.black),\n        ])\n        \n        table.setStyle(style)\n        elements.append(table)\n        \n        # Pied de page avec ROI DU TIR\n        if shooters_sorted:\n            roi = shooters_sorted[0]\n            elements.append(Spacer(1, 10*mm))\n            \n            roi_text = f\"👑 ROI DU TIR: {roi.prenom} {roi.nom} ({roi.societe}) 👑\"\n            roi_style = ParagraphStyle(\n                'ROIStyle',\n                parent=styles['Heading2'],\n                fontSize=14,\n                textColor=colors.HexColor('#FFD700'),\n                alignment=1,\n                fontName='Helvetica-Bold',\n                textColor=colors.HexColor('#003366')\n            )\n            elements.append(Paragraph(roi_text, roi_style))\n        \n        # Générer le PDF\n        doc.build(elements)\n        \n        return output_path\n    \n    def generate_category_ranking(self, category: str = 'Seniors', filename: str = None) -> Path:\n        \"\"\"Génère le classement par catégorie\"\"\"\n        if not filename:\n            timestamp = datetime.now().strftime(\"%Y%m%d_%H%M%S\")\n            filename = f\"02_Classement_Categorie_{category}_{timestamp}.pdf\"\n        \n        output_path = self.output_dir / filename\n        \n        # Récupérer les tireurs de la catégorie\n        shooters = self.db.get_all_shooters()\n        category_shooters = [s for s in shooters if s.categorie == category]\n        \n        if not category_shooters:\n            category_shooters = []\n        \n        # Trier par score\n        category_shooters = sorted(category_shooters, \n            key=lambda x: getattr(x, 'score', 0) or 0, reverse=True)\n        \n        # Créer le document\n        doc = SimpleDocTemplate(\n            str(output_path),\n            pagesize=A4,\n            topMargin=10*mm,\n            bottomMargin=10*mm,\n            leftMargin=10*mm,\n            rightMargin=10*mm,\n            title=f\"Classement Catégorie {category}\"\n        )\n        \n        elements = []\n        styles = getSampleStyleSheet()\n        \n        # Titre\n        title_style = ParagraphStyle(\n            'CustomTitle',\n            parent=styles['Heading1'],\n            fontSize=20,\n            textColor=colors.HexColor('#003366'),\n            spaceAfter=15,\n            alignment=1,\n            fontName='Helvetica-Bold'\n        )\n        \n        elements.append(Paragraph(f\"Classement Catégorie: {category}\", title_style))\n        elements.append(Spacer(1, 10*mm))\n        \n        # Tableau\n        data = [['Rang', 'Nom', 'Prénom', 'Société', 'Score', 'Série', '10s']]\n        \n        for rank, shooter in enumerate(category_shooters, 1):\n            score = getattr(shooter, 'score', 0) or 0\n            serie = getattr(shooter, 'serie', 0) or 0\n            n10 = getattr(shooter, 'n10', 0) or 0\n            \n            data.append([\n                str(rank),\n                shooter.nom,\n                shooter.prenom,\n                shooter.societe or '',\n                f\"{score:.1f}\" if isinstance(score, (int, float)) else str(score),\n                str(serie),\n                str(n10)\n            ])\n        \n        table = Table(data, colWidths=[12*mm, 30*mm, 30*mm, 40*mm, 18*mm, 18*mm, 12*mm])\n        table.setStyle(self._get_table_style())\n        elements.append(table)\n        \n        doc.build(elements)\n        return output_path\n    \n    def generate_society_ranking(self, society: str = 'Tous', filename: str = None) -> Path:\n        \"\"\"Génère le classement par société\"\"\"\n        if not filename:\n            timestamp = datetime.now().strftime(\"%Y%m%d_%H%M%S\")\n            filename = f\"03_Classement_Societe_{society}_{timestamp}.pdf\"\n        \n        output_path = self.output_dir / filename\n        \n        # Récupérer les tireurs de la société\n        shooters = self.db.get_all_shooters()\n        if society != 'Tous':\n            society_shooters = [s for s in shooters if (s.societe or '') == society]\n        else:\n            society_shooters = shooters\n        \n        # Trier par score\n        society_shooters = sorted(society_shooters, \n            key=lambda x: getattr(x, 'score', 0) or 0, reverse=True)\n        \n        # Créer le document\n        doc = SimpleDocTemplate(\n            str(output_path),\n            pagesize=A4,\n            topMargin=10*mm,\n            bottomMargin=10*mm,\n            leftMargin=10*mm,\n            rightMargin=10*mm,\n            title=f\"Classement Société {society}\"\n        )\n        \n        elements = []\n        styles = getSampleStyleSheet()\n        \n        # Titre\n        title_style = ParagraphStyle(\n            'CustomTitle',\n            parent=styles['Heading1'],\n            fontSize=20,\n            textColor=colors.HexColor('#003366'),\n            spaceAfter=15,\n            alignment=1,\n            fontName='Helvetica-Bold'\n        )\n        \n        elements.append(Paragraph(f\"Classement Société: {society}\", title_style))\n        elements.append(Spacer(1, 10*mm))\n        \n        # Tableau avec colonne Prix\n        data = [['Rang', 'Nom', 'Prénom', 'Catégorie', 'Score', 'Série', '10s', 'Prix']]\n        \n        for rank, shooter in enumerate(society_shooters, 1):\n            score = getattr(shooter, 'score', 0) or 0\n            serie = getattr(shooter, 'serie', 0) or 0\n            n10 = getattr(shooter, 'n10', 0) or 0\n            \n            # Déterminer le prix selon le rang\n            if rank == 1:\n                prix = \"100€\"\n            elif rank == 2:\n                prix = \"75€\"\n            elif rank == 3:\n                prix = \"50€\"\n            else:\n                prix = \"\"\n            \n            data.append([\n                str(rank),\n                shooter.nom,\n                shooter.prenom,\n                shooter.categorie,\n                f\"{score:.1f}\" if isinstance(score, (int, float)) else str(score),\n                str(serie),\n                str(n10),\n                prix\n            ])\n        \n        table = Table(data, colWidths=[12*mm, 25*mm, 25*mm, 25*mm, 18*mm, 15*mm, 12*mm, 15*mm])\n        table.setStyle(self._get_table_style())\n        elements.append(table)\n        \n        doc.build(elements)\n        return output_path\n    \n    def generate_summary_report(self, filename: str = None) -> Path:\n        \"\"\"Génère un résumé des classements\"\"\"\n        if not filename:\n            timestamp = datetime.now().strftime(\"%Y%m%d_%H%M%S\")\n            filename = f\"04_Resume_Classements_{timestamp}.pdf\"\n        \n        output_path = self.output_dir / filename\n        \n        shooters = self.db.get_all_shooters()\n        \n        doc = SimpleDocTemplate(\n            str(output_path),\n            pagesize=A4,\n            topMargin=10*mm,\n            bottomMargin=10*mm,\n            leftMargin=10*mm,\n            rightMargin=10*mm,\n            title=\"Résumé des Classements\"\n        )\n        \n        elements = []\n        styles = getSampleStyleSheet()\n        \n        # Titre\n        title_style = ParagraphStyle(\n            'CustomTitle',\n            parent=styles['Heading1'],\n            fontSize=20,\n            textColor=colors.HexColor('#003366'),\n            spaceAfter=15,\n            alignment=1,\n            fontName='Helvetica-Bold'\n        )\n        \n        elements.append(Paragraph(\"Résumé des Classements\", title_style))\n        elements.append(Spacer(1, 10*mm))\n        \n        # Statistiques globales\n        stats_style = ParagraphStyle(\n            'Stats',\n            parent=styles['Normal'],\n            fontSize=11,\n            spaceAfter=10\n        )\n        \n        elements.append(Paragraph(f\"<b>Total de tireurs:</b> {len(shooters)}\", stats_style))\n        elements.append(Spacer(1, 5*mm))\n        \n        # Répartition par catégorie\n        categories = group_by_category(shooters)\n        elements.append(Paragraph(\"<b>Répartition par catégorie:</b>\", stats_style))\n        for cat, shooters_cat in sorted(categories.items()):\n            elements.append(Paragraph(f\"• {cat}: {len(shooters_cat)} tireurs\", stats_style))\n        \n        elements.append(Spacer(1, 10*mm))\n        \n        # Répartition par société\n        societies = group_by_society(shooters)\n        elements.append(Paragraph(\"<b>Répartition par société:</b>\", stats_style))\n        for soc, shooters_soc in sorted(societies.items()):\n            elements.append(Paragraph(f\"• {soc}: {len(shooters_soc)} tireurs\", stats_style))\n        \n        doc.build(elements)\n        return output_path\n    \n    def generate_all_stand_sheets(self) -> List[Path]:\n        \"\"\"Génère les feuilles de stand (C6) pour tous les tireurs\"\"\"\n        shooters = self.db.get_all_shooters()\n        paths = []\n        \n        for shooter in shooters:\n            path = self._generate_stand_sheet(shooter)\n            paths.append(path)\n        \n        return paths\n    \n    def _generate_stand_sheet(self, shooter: Shooter, filename: str = None) -> Path:\n        \"\"\"Génère une feuille de stand (C6) pour un tireur\"\"\"\n        if not filename:\n            filename = f\"Stand_{shooter.prenom}_{shooter.nom}_{datetime.now().strftime('%Y%m%d')}.pdf\"\n        \n        output_path = self.output_dir / filename\n        \n        # C6 en mm: 114 x 162\n        c6_width = 114*mm\n        c6_height = 162*mm\n        \n        c = canvas.Canvas(str(output_path), pagesize=(c6_width, c6_height))\n        \n        # Cadre\n        c.setLineWidth(1)\n        c.rect(5*mm, 5*mm, c6_width-10*mm, c6_height-10*mm)\n        \n        # En-tête\n        c.setFont(\"Helvetica-Bold\", 14)\n        c.drawString(10*mm, c6_height-15*mm, \"FEUILLE DE TIR\")\n        \n        # Informations du tireur\n        c.setFont(\"Helvetica-Bold\", 11)\n        c.drawString(10*mm, c6_height-30*mm, f\"Nom: {shooter.nom}\")\n        c.setFont(\"Helvetica\", 11)\n        c.drawString(10*mm, c6_height-38*mm, f\"Prénom: {shooter.prenom}\")\n        c.drawString(10*mm, c6_height-46*mm, f\"Société: {shooter.societe or 'N/A'}\")\n        c.drawString(10*mm, c6_height-54*mm, f\"Arme: {shooter.arme or 'N/A'}\")\n        c.drawString(10*mm, c6_height-62*mm, f\"Catégorie: {shooter.categorie}\")\n        \n        # Tableau pour les résultats\n        c.setFont(\"Helvetica-Bold\", 9)\n        c.drawString(10*mm, c6_height-75*mm, \"Résultats:\")\n        \n        # Lignes de saisie\n        y = c6_height - 85*mm\n        for i in range(5):\n            c.setFont(\"Helvetica\", 9)\n            c.drawString(10*mm, y, f\"Tir {i+1}: Points ___  Série ___  10s ___\")\n            y -= 10*mm\n        \n        # Signature\n        c.setFont(\"Helvetica-Bold\", 9)\n        c.drawString(10*mm, 10*mm, \"Signature du tireur: ________________\")\n        \n        c.save()\n        \n        return output_path\n    \n    def _get_table_style(self):\n        \"\"\"Retourne le style standard pour les tableaux\"\"\"\n        return TableStyle([\n            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#003366')),\n            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),\n            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),\n            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),\n            ('FONTSIZE', (0, 0), (-1, 0), 9),\n            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),\n            ('ALIGN', (0, 1), (-1, -1), 'CENTER'),\n            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),\n            ('FONTSIZE', (0, 1), (-1, -1), 8),\n            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),\n            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F5F5F5')])\n        ])\n