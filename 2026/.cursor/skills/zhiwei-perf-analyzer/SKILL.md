---
name: zhiwei-perf-analyzer
description: Analyze and optimize performance in ZhiWei project. Use when user asks about performance, slow operations, memory usage, or needs to optimize code for speed/efficiency.
---

# ZhiWei Performance Analyzer

Identify and fix performance bottlenecks in the ZhiWei project.

## Known Performance Issues

### 1. Large List Memory (P2)

**Files affected:**
- `converter/scan_convert_candidates.py` - os.walk results loaded into single list
- `converter/collect_index.py` - unique_records fully in memory
- `converter/chromadb_docs.py` - all docs loaded before batching

**Solution:** Convert to generators (yield) for streaming processing.

### 2. Parallel Task Submission (P2)

**File:** `converter/batch_parallel.py`

**Issue:** All tasks submitted at once, creating many pending futures.

**Solution:** Batch submission with limited pending count.

### 3. Concurrency Model

**Current:** ThreadPoolExecutor (GIL limited)

**Note:** Kept for Office COM compatibility. Do NOT switch to multiprocessing without considering COM threading requirements.

## Performance Checklist

When optimizing, verify:
- [ ] Memory usage doesn't spike during large file processing
- [ ] Progress callbacks don't block the main thread
- [ ] Checkpoint intervals balance safety vs performance
- [ ] Generator patterns used where appropriate

## Profiling Commands

```bash
# Profile a specific function
python -m cProfile -s time office_converter.py --source "test_dir" --run-mode convert_only

# Memory profiling (requires memory_profiler)
python -m memory_profiler office_converter.py --source "test_dir"
```

## Metrics to Track

| Metric | Target | Current |
|--------|--------|---------|
| Test suite time | < 15s | ~10s (394 tests) |
| Large batch (1000 files) | < 5min | TBD |
| Memory per 1000 files | < 500MB | TBD |

## Optimization Priority

1. **P2** - Generator conversion for scan/collect operations
2. **P2** - Batch submission in parallel processing
3. **P3** - Path unification to pathlib (incremental)