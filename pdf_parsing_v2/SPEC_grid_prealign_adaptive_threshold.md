## Goal
Design a production follow-up for `grid_matcher.adapt_by_cell_assignment()` so that affine scale threshold is no longer a fixed magic value after proposed-bbox pre-alignment.

## Current temporary policy
- Production now uses proposed-bbox pre-alignment before matching.
- The legacy path keeps `scale_threshold = 0.05`.
- Only the proposed-bbox path uses a temporary reduced threshold `0.04`.
- This is intentionally a guardrailed interim solution, not the final design.

## Why adaptive threshold is needed
- Benchmark data shows that reduced threshold helps once `proposed_bbox` is reliable.
- The same fixed threshold is not ideal across all deformation regimes.
- If threshold is reduced when pre-alignment is weak, noisy `d_all` can trigger false global scaling and distort the whole stamp geometry.

## Confidence signals
Use only signals already available inside pre-alignment helpers or cheap to derive there.

### Extent-probe quality
- `extent_probe_top_candidate_count_strong`
- `extent_probe_left_candidate_count_strong`
- `extent_probe_virtual_top_candidate_count`
- `extent_probe_virtual_left_candidate_count`
- whether `virtual_outer_top` and `virtual_outer_left` were inferred at all
- whether inferred edges came from `*_virtual` or only from raw edge candidates

### Boundary span quality
- span ratio of chosen top/left/right/bottom boundary-like lines against template bbox width/height
- whether chosen left/right candidates cover the full expected vertical span
- whether proposed bbox keeps stable right/bottom edges from the base stamp bbox for `bottom_right`

### Proposed-bbox stability
- `abs(proposed_bbox.x0 - base_bbox.x0) / base_bbox.width`
- `abs(proposed_bbox.y0 - base_bbox.y0) / base_bbox.height`
- resulting width/height ratio between proposed bbox and base bbox
- whether filtering by proposed bbox still keeps a dense detected pool

### Post-bbox geometry quality
- clipped cell count inside effective bbox
- rectangularity score after effective-bbox filtering
- agreement of `sx_raw` and `sy_raw` with the inferred bbox deltas

## Proposed policy
Compute a coarse confidence tier before deciding `use_scale`.

### High confidence
- virtual top and virtual left exist
- strong candidate counts are non-trivial
- boundary span ratios are near-full-span
- proposed bbox keeps a healthy detected pool and good rectangularity

Policy:
- allow reduced threshold in the `0.04..0.045` class

### Medium confidence
- only one virtual edge exists, or strong candidates are mixed with fallback ones
- proposed bbox is plausible but not strongly confirmed

Policy:
- use a midpoint threshold around `0.045..0.05`

### Low confidence
- no virtual edges, weak candidate pools, poor span ratios, or rectangularity degrades after bbox filtering

Policy:
- fall back to legacy `0.05`
- optionally skip proposed-bbox-derived scale altogether while still allowing bbox filtering if that remains useful

## Implementation shape
- Keep the adaptive decision inside `grid_matcher.py`; do not spread it across benchmark tooling.
- Add one helper that converts pre-alignment diagnostics into:
  - `confidence_tier`
  - `scale_threshold`
  - human-readable reason string for `warnings`
- Preserve the existing fallback behavior when extent probe fails or returns an empty detected pool.

## Verification
- Re-run the synthetic benchmark on:
  - the former `grid256` set
  - the denser `90..110` set
- Compare against:
  - production temporary `0.04`
  - benchmark reference `proposed_bbox_y20_scale45`
- Spot-check several ordinary non-deformed PDFs to ensure the adaptive path does not introduce false scaling regressions.
