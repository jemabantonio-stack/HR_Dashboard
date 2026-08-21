/* ==========================================================================
 * Hog Ranch Dashboard — published data auto-loader
 * --------------------------------------------------------------------------
 * Drop-in module. The ONLY change required in index.html is one line, placed
 * immediately before </body> and AFTER the main inline <script>:
 *
 *     <script src="autoload.js"></script>
 *
 * What it does:
 *   1. Reads manifest.json (same folder as index.html).
 *   2. Fetches every dataset the manifest lists and renders it — no manual
 *      file selection required.
 *   3. Collapses the existing Excel / Pads / Tenement file pickers into a
 *      "Load custom data" disclosure, kept as a fallback for testing.
 *
 * It reuses the functions and state declared by the inline script
 * (parseWorkbook, renderShapeLayer, loadedShapeGeoJson, messages, ...). Those
 * are top-level declarations in a classic script, so they are visible here
 * because this file executes afterwards. Nothing in index.html needs editing.
 * ========================================================================== */

(function () {
  'use strict';

  var MANIFEST_PATH = 'manifest.json';

  var panelStyles = [
    '.controls-row.manual-row { display: block; }',
    '.manual-load { width: max-content; max-width: 100%; margin-left: auto; box-sizing: border-box;',
    '  border-radius: 14px; background: rgba(11, 52, 70, 0.22);',
    '  border: 1px solid rgba(160, 166, 120, 0.25); }',
    '.manual-load > summary { display: flex; align-items: center; gap: 8px;',
    '  min-height: 36px; padding: 0 12px; list-style: none; cursor: pointer;',
    '  color: #A0A678; font-size: 0.8rem; font-weight: 600; white-space: nowrap; }',
    '.manual-load > summary::-webkit-details-marker { display: none; }',
    '.manual-load > summary::before { content: "\\25B8"; font-size: 0.7rem; }',
    '.manual-load[open] > summary::before { content: "\\25BE"; }',
    '.manual-load .data-source-note { margin-left: auto; font-weight: 500;',
    '  color: rgba(245, 248, 255, 0.6); }',
    '.manual-load-body { display: flex; align-items: stretch; gap: 8px;',
    '  padding: 0 8px 8px 8px; }',
    '.manual-load .control-capsule { width: 240px; }',
    '.manual-load .control-capsule input[type="file"] { width: 148px;',
    '  min-width: 148px; max-width: 148px; }',
    '.manual-load .reload-button { align-self: center; height: 28px; padding: 0 10px;',
    '  border: 1px solid rgba(160, 166, 120, 0.25); border-radius: 10px;',
    '  background: rgba(67, 51, 28, 0.25); color: #f5f8ff; font-weight: 700;',
    '  font-size: 0.8rem; cursor: pointer; white-space: nowrap; }',
    '.manual-load .reload-button:hover { background: rgba(67, 51, 28, 0.4); }',
    '.manual-load .reload-button[disabled] { opacity: 0.5; cursor: default; }'
  ].join('\n');

  var statusBar = document.getElementById('messages');
  var manualPanel = null;
  var sourceNote = null;
  var reloadButton = null;
  var autoLoadInFlight = false;

  function setStatus(text) {
    if (statusBar) statusBar.textContent = text;
  }

  function setNote(text) {
    if (sourceNote) sourceNote.textContent = text;
  }

  /* ---------------------------------------------------------------------
   * UI: fold the existing file pickers into a collapsed disclosure.
   * Purely cosmetic — if the markup ever changes and this bails out, the
   * auto-loading below still works and the pickers simply stay visible.
   * ------------------------------------------------------------------ */
  function buildManualPanel() {
    var fileInput = document.getElementById('fileInput');
    if (!fileInput || !fileInput.closest) return;

    var capsule = fileInput.closest('.control-capsule');
    var row = capsule ? capsule.closest('.controls-row') : null;
    if (!row || row.querySelector('.manual-load')) return;

    var style = document.createElement('style');
    style.textContent = panelStyles;
    document.head.appendChild(style);

    row.classList.add('manual-row');

    var details = document.createElement('details');
    details.className = 'manual-load';
    details.id = 'manualLoadPanel';

    var summary = document.createElement('summary');
    summary.appendChild(document.createTextNode('Load custom data'));

    var note = document.createElement('span');
    note.className = 'data-source-note';
    note.id = 'dataSourceNote';
    note.textContent = 'Loading…';
    summary.appendChild(note);

    var body = document.createElement('div');
    body.className = 'manual-load-body';
    while (row.firstChild) {
      body.appendChild(row.firstChild);
    }

    var reload = document.createElement('button');
    reload.type = 'button';
    reload.className = 'reload-button';
    reload.id = 'reloadDataButton';
    reload.textContent = 'Reload published';
    body.appendChild(reload);

    details.appendChild(summary);
    details.appendChild(body);
    row.appendChild(details);

    manualPanel = details;
    sourceNote = note;
    reloadButton = reload;

    reload.addEventListener('click', function () {
      setStatus('Reloading published data…');
      autoLoadPublishedData();
    });

    ['fileInput', 'shapefileInput', 'tenementInput'].forEach(function (id) {
      var input = document.getElementById(id);
      if (!input) return;
      input.addEventListener('change', function () {
        if (input.files && input.files.length) setNote('Custom data loaded');
      });
    });
  }

  /* ---------------------------------------------------------------------
   * Fetch helpers
   * ------------------------------------------------------------------ */

  /*
   * cache: 'no-cache' forces a conditional request on every load. GitHub Pages
   * answers 304 when the file is unchanged, so this costs almost nothing, but
   * a freshly committed file is picked up immediately instead of being masked
   * by the CDN's default cache window. Preferred over a ?v=timestamp param,
   * which would force a full re-download of the ~500 KB workbook every visit.
   */
  function fetchPublished(path) {
    return fetch(path, { cache: 'no-cache' }).then(function (response) {
      if (!response.ok) {
        throw new Error('HTTP ' + response.status + ' for ' + path);
      }
      return response;
    });
  }

  function datasetCandidates(entry) {
    if (!entry) return [];
    var list = [];
    if (typeof entry === 'string') {
      list.push(entry);
    } else {
      if (entry.path) list.push(entry.path);
      if (Array.isArray(entry.fallbacks)) list = list.concat(entry.fallbacks);
    }
    return list.filter(function (value) {
      return typeof value === 'string' && value.trim().length > 0;
    });
  }

  function datasetLabel(entry, fallbackLabel) {
    if (entry && typeof entry === 'object' && entry.name) return entry.name;
    return fallbackLabel;
  }

  function resolveDataset(entry) {
    var candidates = datasetCandidates(entry);
    if (!candidates.length) {
      return Promise.reject(new Error('no path configured in manifest.json'));
    }

    var index = 0;
    var lastError = null;

    function attempt() {
      if (index >= candidates.length) {
        return Promise.reject(lastError || new Error('unable to resolve dataset'));
      }
      var candidate = candidates[index++];
      return fetchPublished(candidate).then(
        function (response) {
          return { response: response, path: candidate };
        },
        function (error) {
          lastError = error;
          return attempt();
        }
      );
    }

    return attempt();
  }

  /* ---------------------------------------------------------------------
   * Rendering — mirrors the inline script's manual-load path so fetched and
   * hand-picked data go through identical normalisation.
   * ------------------------------------------------------------------ */
  function applySpatialData(raw, target, sourceLabel, fitToLayer) {
    var isTenement = target === 'tenement';
    var featureCollection = toFeatureCollection(raw);

    if (!featureCollection || !featureCollection.features.length) {
      throw new Error('no features found in ' + sourceLabel);
    }

    var normalized = normalizeFeatureCollectionCoordinates(featureCollection);

    if (isTenement) {
      loadedTenementGeoJson = normalized;
      renderTenementLayer(loadedTenementGeoJson, fitToLayer);
    } else {
      loadedShapeGeoJson = normalized;
      renderShapeLayer(loadedShapeGeoJson, fitToLayer);
      updateStatusSummary();
      if (typeof currentSelectedPoint !== 'undefined' && currentSelectedPoint) {
        updateDetails(currentSelectedPoint);
      }
    }

    return normalized.features.length;
  }

  function autoLoadSpatial(entry, target, fallbackLabel) {
    return resolveDataset(entry).then(function (result) {
      var path = result.path;
      var isZip = path.toLowerCase().indexOf('.zip', path.length - 4) !== -1;

      if (isZip) {
        if (typeof shp === 'undefined') {
          throw new Error('shapefile parser unavailable; publish GeoJSON instead of .zip');
        }
        return result.response.arrayBuffer().then(function (buffer) {
          return shp(buffer);
        }).then(function (raw) {
          return applySpatialData(raw, target, datasetLabel(entry, fallbackLabel), false);
        });
      }

      return result.response.json().then(function (raw) {
        return applySpatialData(raw, target, datasetLabel(entry, fallbackLabel), false);
      });
    });
  }

  function autoLoadWorkbook(entry) {
    if (typeof XLSX === 'undefined') {
      return Promise.reject(new Error('XLSX library failed to load'));
    }

    return resolveDataset(entry).then(function (result) {
      // Excel is binary: it must be read as an ArrayBuffer, never as text.
      return result.response.arrayBuffer().then(function (buffer) {
        var workbook = parseArrayBufferToWorkbook(buffer);
        parseWorkbook(workbook);
        if (typeof loadedRows === 'undefined' || !loadedRows || !loadedRows.length) {
          throw new Error('no drillhole rows parsed from ' + result.path);
        }
        return loadedRows.length;
      });
    });
  }

  /* ---------------------------------------------------------------------
   * Orchestration
   * ------------------------------------------------------------------ */
  function autoLoadPublishedData() {
    if (autoLoadInFlight) return Promise.resolve();
    autoLoadInFlight = true;
    if (reloadButton) reloadButton.disabled = true;

    var loaded = [];
    var failed = [];
    var manifest = null;

    function finish() {
      autoLoadInFlight = false;
      if (reloadButton) reloadButton.disabled = false;
    }

    return fetchPublished(MANIFEST_PATH)
      .then(function (response) { return response.json(); })
      .then(function (parsed) {
        manifest = parsed || {};
        var datasets = manifest.datasets || {};
        var chain = Promise.resolve();

        // Order matters: the two spatial layers render without auto-fit so the
        // final map extent is driven by the drillhole collars in the workbook.
        [
          { key: 'tenement', target: 'tenement', label: 'Tenement' },
          { key: 'pads', target: 'pads', label: 'Drill pads' }
        ].forEach(function (job) {
          if (!datasets[job.key]) return;
          chain = chain.then(function () {
            return autoLoadSpatial(datasets[job.key], job.target, job.label).then(
              function () { loaded.push(datasetLabel(datasets[job.key], job.label)); },
              function (error) {
                console.error(job.label + ' auto-load failed:', error);
                failed.push(job.label + ' (' + error.message + ')');
              }
            );
          });
        });

        if (datasets.workbook) {
          chain = chain.then(function () {
            return autoLoadWorkbook(datasets.workbook).then(
              function () { loaded.push(datasetLabel(datasets.workbook, 'Workbook')); },
              function (error) {
                console.error('Workbook auto-load failed:', error);
                failed.push('Workbook (' + error.message + ')');
              }
            );
          });
        }

        return chain;
      })
      .then(function () {
        var stamp = manifest && manifest.updated ? ' (data as of ' + manifest.updated + ')' : '';

        if (!loaded.length && !failed.length) {
          setStatus('manifest.json lists no datasets. Open "Load custom data" to select files manually.');
          setNote('No datasets');
          if (manualPanel) manualPanel.open = true;
        } else if (failed.length) {
          setStatus('Loaded: ' + (loaded.join(', ') || 'nothing') +
            '. Failed: ' + failed.join('; ') + '. See browser console for detail.');
          setNote(loaded.length + ' of ' + (loaded.length + failed.length) + ' loaded');
          if (manualPanel) manualPanel.open = true;
        } else {
          setStatus('Loaded published data: ' + loaded.join(', ') + stamp + '.');
          setNote('Published data' + stamp);
        }
        finish();
      })
      .catch(function (error) {
        console.error('Manifest load failed:', error);
        setStatus('Could not read manifest.json (' + error.message +
          '). Open "Load custom data" to select files manually.');
        setNote('Manifest missing');
        if (manualPanel) manualPanel.open = true;
        finish();
      });
  }

  buildManualPanel();
  setStatus('Loading published data…');
  autoLoadPublishedData();
})();
