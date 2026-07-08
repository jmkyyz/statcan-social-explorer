#!/usr/bin/env python3
"""Pre-build the gzipped catalog cache at deploy time.

Run from render.yaml's buildCommand (after pip install) so a fresh instance
serves /api/catalog and /api/catalog-rows instantly from catalog_cache/
instead of streaming the 140k-row Vectors.xlsx on first request — which takes
minutes on Render's 0.5 vCPU and times out behind Cloudflare (~100s).

Also fine to run locally: the cache dir is gitignored, and the server
re-validates it against Vectors.xlsx by content hash before trusting it.
"""
import time

import proxy

start = time.time()
info = proxy.write_disk_cache()
print(f"catalog cache built in {time.time() - start:.1f}s → {proxy.CACHE_DIR}")
print(f"  {info['categories']} categories, {info['series']} series, "
      f"rows gz {info['rows_gz_bytes'] / 1e6:.1f}MB")
