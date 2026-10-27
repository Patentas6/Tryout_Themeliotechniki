# -*- coding: utf-8 -*-
"""
count_piles_opencv.py — Καταμέτρηση Πασσάλων ΠΣ (Ø800mm) & ΠΣ1 (Ø600mm) με OpenCV & PyMuPDF
Έργο: Atlantic Forest — Plot 1 (Παρεκκλησιά, Λεμεσός)
Σχέδιο: PILING DRAWINGS - PLOT 1.pdf (Φύλλο 00Σ)
"""
import os
import sys
import fitz
import cv2
import numpy as np

sys.stdout.reconfigure(encoding="utf-8")

DIR = os.path.dirname(os.path.abspath(__file__))
PDF_PATH = os.path.join(DIR, "PILING DRAWINGS - PLOT 1.pdf")
OUT_IMG = os.path.join(DIR, "piles_detected_opencv.png")


def main():
    doc = fitz.open(PDF_PATH)
    page = doc[0]
    mat = fitz.Matrix(3.0, 3.0)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
    img_bgr = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

    b, g, r = cv2.split(img_bgr)
    red_mask = ((r > 180) & (g < 80) & (b < 80)).astype(np.uint8) * 255
    blue_mask = ((b > 200) & (r < 50) & (g < 50)).astype(np.uint8) * 255

    # Εξαίρεση των δύο υπομνημάτων λεπτομέρειας στο κάτω μέρος του σχεδίου
    red_mask[int(880 * 3):, :int(380 * 3)] = 0
    blue_mask[int(1020 * 3):, int(430 * 3):int(780 * 3)] = 0

    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (9, 9))
    red_closed = cv2.morphologyEx(red_mask, cv2.MORPH_CLOSE, kernel)
    blue_closed = cv2.morphologyEx(blue_mask, cv2.MORPH_CLOSE, kernel)

    num_r, _, stats_r, centroids_r = cv2.connectedComponentsWithStats(red_closed)
    ps_piles = []
    for i in range(1, num_r):
        w_box, h_box = stats_r[i, cv2.CC_STAT_WIDTH], stats_r[i, cv2.CC_STAT_HEIGHT]
        if 50 < w_box < 90 and 50 < h_box < 90:
            ps_piles.append((centroids_r[i][0], centroids_r[i][1]))

    num_b, _, stats_b, centroids_b = cv2.connectedComponentsWithStats(blue_closed)
    ps1_piles = []
    for i in range(1, num_b):
        w_box, h_box = stats_b[i, cv2.CC_STAT_WIDTH], stats_b[i, cv2.CC_STAT_HEIGHT]
        if 40 < w_box < 75 and 40 < h_box < 75:
            ps1_piles.append((centroids_b[i][0], centroids_b[i][1]))

    ann = img_bgr.copy()
    for idx, (cx, cy) in enumerate(ps_piles, 1):
        ix, iy = int(round(cx)), int(round(cy))
        cv2.circle(ann, (ix, iy), 45, (0, 0, 255), 4)
        cv2.putText(ann, f"PS-{idx}", (ix + 48, iy + 10), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 0, 200), 3)

    for idx, (cx, cy) in enumerate(ps1_piles, 1):
        ix, iy = int(round(cx)), int(round(cy))
        cv2.circle(ann, (ix, iy), 38, (255, 0, 0), 4)
        cv2.putText(ann, f"PS1-{idx}", (ix + 42, iy + 10), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (200, 0, 0), 3)

    cv2.imwrite(OUT_IMG, ann)
    print(f"Πάσσαλοι ΠΣ (Ø800mm, L=10m): {len(ps_piles)} τεμ.")
    print(f"Πάσσαλοι ΠΣ1 (Ø600mm, L=5m): {len(ps1_piles)} τεμ.")
    print(f"Σύνολο Πασσάλων Έργου: {len(ps_piles) + len(ps1_piles)} τεμ.")
    print(f"Αποθηκεύτηκε εικόνα επαλήθευσης: {OUT_IMG}")


if __name__ == "__main__":
    main()
