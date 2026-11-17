#!/usr/bin/env python3
"""
compare_images.py — Compare two images from URLs.

Usage:
    python compare_images.py <url1> <url2>
"""

import sys
import io
import hashlib
import requests
import numpy as np
from PIL import Image, ImageChops
from skimage.metrics import structural_similarity as ssim
import cv2
from imagehash import phash, dhash, average_hash

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; ImageCompareBot/1.0)"}


def fetch_image(url: str) -> bytes:
    """Download image bytes from URL."""
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    if not r.headers.get("content-type", "").startswith("image"):
        # Some servers mislabel; still try to open
        pass
    return r.content


def load_image(raw: bytes) -> Image.Image:
    return Image.open(io.BytesIO(raw)).convert("RGB")


# ---------- Comparison methods ----------

def byte_hash(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def pixel_identical(img1: Image.Image, img2: Image.Image) -> bool:
    if img1.size != img2.size:
        return False
    diff = ImageChops.difference(img1, img2)
    return diff.getbbox() is None


def phash_distance(img1: Image.Image, img2: Image.Image) -> int:
    return phash(img1) - phash(img2)


def dhash_distance(img1: Image.Image, img2: Image.Image) -> int:
    return dhash(img1) - dhash(img2)

def ahash_distance(img1: Image.Image, img2: Image.Image) -> int:
    return average_hash(img1) - average_hash(img2)

def ssim_score(img1: Image.Image, img2: Image.Image) -> float:
    """SSIM requires same size — resize img2 to img1's size."""
    a = np.array(img1.convert("L"))
    b = np.array(img2.convert("L").resize(img1.size))
    score, _ = ssim(a, b, full=True)
    return float(score)


def orb_match_score(img1: Image.Image, img2: Image.Image) -> float:
    """Feature-match score. Returns fraction of good matches."""
    g1 = cv2.cvtColor(np.array(img1), cv2.COLOR_RGB2GRAY)
    g2 = cv2.cvtColor(np.array(img2), cv2.COLOR_RGB2GRAY)

    orb = cv2.ORB_create(nfeatures=2000)
    kp1, des1 = orb.detectAndCompute(g1, None)
    kp2, des2 = orb.detectAndCompute(g2, None)

    if des1 is None or des2 is None:
        return 0.0

    bf = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
    matches = bf.match(des1, des2)
    if not matches:
        return 0.0

    # Ratio of good matches to smaller keypoint set
    good = [m for m in matches if m.distance < 50]
    return len(good) / min(len(kp1), len(kp2))


# ---------- Verdict logic ----------

def verdict(results: dict) -> str:
    if results["pixel_identical"]:
        return "IDENTICAL (pixel-perfect)"

    ph = results["phash_distance"]
    dh = results["dhash_distance"]
    ah = results["ahash_distance"]
    s = results["ssim"]
    orb = results["orb_score"]

    if ph == 0 and dh <= 2 and ah <= 2:
        return "VISUALLY IDENTICAL (same image, re-encoded)"
    if s >= 0.98 and ph <= 3:
        return "NEARLY IDENTICAL (very minor differences)"
    if s >= 0.90 or (ph <= 8 and orb >= 0.15):
        return "VERY SIMILAR (likely same image, slight edits/resize)"
    if s >= 0.75 or orb >= 0.25:
        return "SIMILAR (possibly related — same subject or content)"
    if orb >= 0.10:
        return "WEAKLY RELATED (may share some content)"
    return "DIFFERENT"


# ---------- Main ----------

def compare(url1: str, url2: str):
    print(f"Fetching:\n  1) {url1}\n  2) {url2}\n")
    raw1 = fetch_image(url1)
    raw2 = fetch_image(url2)
    img1 = load_image(raw1)
    img2 = load_image(raw2)

    print(f"Sizes: {img1.size}  vs  {img2.size}\n")

    results = {
        "sha256_match":      byte_hash(raw1) == byte_hash(raw2),
        "pixel_identical":   pixel_identical(img1, img2),
        "phash_distance":    phash_distance(img1, img2),
        "dhash_distance":    dhash_distance(img1, img2),
        "ahash_distance":    ahash_distance(img1, img2),
        "ssim":              ssim_score(img1, img2),
        "orb_score":         orb_match_score(img1, img2),
    }

    print("=== Results ===")
    print(f"  SHA-256 match (byte-identical) : {results['sha256_match']}")
    print(f"  Pixel-identical                : {results['pixel_identical']}")
    print(f"  pHash distance                 : {results['phash_distance']}  (0=identical, ≤10 similar)")
    print(f"  dHash distance                 : {results['dhash_distance']}")
    print(f"  aHash distance                 : {results['ahash_distance']}")
    print(f"  SSIM                           : {results['ssim']:.4f}  (1.0=identical)")
    print(f"  ORB feature-match ratio        : {results['orb_score']:.4f}")
    print()
    print(f"VERDICT → {verdict(results)}")
    return results


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python compare_images.py <url1> <url2>")
        sys.exit(1)
    compare(sys.argv[1], sys.argv[2])