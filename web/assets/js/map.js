(function () {
  "use strict";

  const data = window.VOYAGE_DATA;
  const OSRM_SERVERS = [
    "https://routing.openstreetmap.de/routed-foot/route/v1/foot",
    "https://router.project-osrm.org/route/v1/foot",
  ];
  const ROUTE_DELAY_MS = 100;
  // Nombre de requetes OSRM simultanees (reste courtois envers les serveurs publics).
  const ROUTE_CONCURRENCY = 3;
  // Palette des trajets. La palette des jours (DAY_COLORS) vit dans
  // scripts/excel_utils.py et arrive via le champ "couleur" du JSON.
  const ROUTE_COLORS = [
    "#e74c3c",
    "#27ae60",
    "#3498db",
    "#9b59b6",
    "#e67e22",
    "#1abc9c",
    "#f39c12",
    "#2c3e50",
    "#d35400",
    "#8e44ad",
    "#16a085",
    "#c0392b",
    "#2980b9",
    "#7f8c8d",
    "#d81b60",
    "#5e35b1",
    "#00897b",
    "#6d4c41",
    "#455a64",
    "#ef6c00",
  ];
  const DEFAULT_DAY_COLOR = "#2c3e50";

  if (!data || !data.points) {
    console.error("VOYAGE_DATA manquant");
    return;
  }

  /* ---------- Carte ---------- */

  const map = L.map("map", { zoomControl: true });
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
    maxZoom: 19,
  }).addTo(map);

  const layerGroups = {};
  const routeLayer = L.layerGroup().addTo(map);
  const markers = [];
  const markersByPointId = new Map();
  const routeCache = new Map();
  const segmentInputs = new Map();
  const segmentMetaEls = new Map();
  const segmentsById = new Map();
  const dayToggleButtons = new Map();
  const dayTotalEls = new Map();
  const dayCheckboxes = new Map();
  let routeRequestId = 0;
  let allSegments = [];

  const escapeDiv = document.createElement("div");

  function escapeHtml(text) {
    if (text == null) return "";
    escapeDiv.textContent = String(text);
    return escapeDiv.innerHTML;
  }

  function markerLabel(point) {
    if (point.ordre_label) return String(point.ordre_label);
    if (point.jour != null && point.visite != null) {
      return point.jour + "." + point.visite;
    }
    if (point.ordre != null) return String(point.ordre);
    return "?";
  }

  function shortName(name) {
    const text = String(name || "");
    return text.length > 28 ? text.slice(0, 25) + "…" : text;
  }

  function formatDistance(meters) {
    if (meters == null) return "";
    if (meters < 950) return Math.round(meters) + " m";
    return (meters / 1000).toFixed(1).replace(".", ",") + " km";
  }

  function formatDuration(seconds) {
    if (seconds == null) return "";
    const minutes = Math.round(seconds / 60);
    if (minutes < 60) return minutes + " min";
    const hours = Math.floor(minutes / 60);
    const rest = minutes % 60;
    return rest > 0 ? hours + " h " + rest + " min" : hours + " h";
  }

  /* ---------- Organisation des points par jour ---------- */

  const allJours = (data.jours || [])
    .map(function (jour) { return String(jour); })
    .sort(function (a, b) { return Number(a) - Number(b); });

  const pointsByDay = {};
  data.points.forEach(function (point) {
    const jourKey = String(point.jour);
    if (!pointsByDay[jourKey]) pointsByDay[jourKey] = [];
    pointsByDay[jourKey].push(point);
  });
  Object.keys(pointsByDay).forEach(function (jourKey) {
    pointsByDay[jourKey].sort(function (a, b) {
      return a.visite - b.visite || String(a.id).localeCompare(String(b.id));
    });
  });

  const dayColors = {};
  data.points.forEach(function (point) {
    const jourKey = String(point.jour);
    if (!dayColors[jourKey] && point.couleur) dayColors[jourKey] = point.couleur;
  });

  function dayColor(jourKey) {
    return dayColors[jourKey] || DEFAULT_DAY_COLOR;
  }

  // Visite precedente/suivante du meme jour, pour la navigation dans les popups.
  const neighborsById = new Map();
  Object.keys(pointsByDay).forEach(function (jourKey) {
    const points = pointsByDay[jourKey];
    points.forEach(function (point, i) {
      neighborsById.set(point.id, {
        prev: i > 0 ? points[i - 1] : null,
        next: i < points.length - 1 ? points[i + 1] : null,
      });
    });
  });

  /* ---------- Popups et marqueurs ---------- */

  function buildPopup(point) {
    const p = point.popup || {};
    const parts = [
      `<div class="popup-content">`,
      `<span class="badge">Jour ${escapeHtml(point.jour)} — visite ${escapeHtml(point.visite)} (${escapeHtml(markerLabel(point))})</span>`,
      `<h3>${escapeHtml(point.nom)}</h3>`,
    ];

    if (p.action || p.type) {
      parts.push(`<p class="meta"><strong>${escapeHtml(p.action || "")}</strong> ${escapeHtml(p.type || "")}</p>`);
    }
    if (p.ouverture || p.fermeture) {
      parts.push(`<p class="meta">Horaires : ${escapeHtml(p.ouverture || "?")} – ${escapeHtml(p.fermeture || "?")}</p>`);
    }
    if (p.prix != null && p.prix !== "") {
      parts.push(`<p class="meta">Prix : ${escapeHtml(p.prix)} €</p>`);
    }
    if (p.billet) {
      parts.push(`<p class="meta">Billet : ${escapeHtml(p.billet)}</p>`);
    }
    if (p.city_card != null && p.city_card !== "") {
      parts.push(`<p class="meta">City Card : ${escapeHtml(p.city_card)}</p>`);
    }
    if (p.remarque) {
      parts.push(`<p class="meta">${escapeHtml(p.remarque)}</p>`);
    }
    if (point.lien) {
      parts.push(`<p class="meta"><a href="${escapeHtml(point.lien)}" target="_blank" rel="noopener">Site web</a></p>`);
    }

    const neighbors = neighborsById.get(point.id) || {};
    if (neighbors.prev || neighbors.next) {
      parts.push(`<div class="popup-nav">`);
      if (neighbors.prev) {
        parts.push(
          `<button type="button" class="popup-nav-btn" data-goto="${escapeHtml(neighbors.prev.id)}" ` +
          `title="${escapeHtml(neighbors.prev.nom)}">← ${escapeHtml(markerLabel(neighbors.prev))}</button>`
        );
      } else {
        parts.push(`<span></span>`);
      }
      if (neighbors.next) {
        parts.push(
          `<button type="button" class="popup-nav-btn" data-goto="${escapeHtml(neighbors.next.id)}" ` +
          `title="${escapeHtml(neighbors.next.nom)}">${escapeHtml(markerLabel(neighbors.next))} →</button>`
        );
      }
      parts.push(`</div>`);
    }

    parts.push(`</div>`);
    return parts.join("");
  }

  function createNumberIcon(point) {
    const label = markerLabel(point);
    const text = escapeHtml(label);
    const ariaLabel = escapeHtml(label + " — " + point.nom);
    const width = Math.max(28, String(label).length * 9 + 14);
    const half = Math.round(width / 2);
    return L.divIcon({
      className: "numbered-marker",
      html: `<div class="marker-number" role="img" aria-label="${ariaLabel}" style="background:${point.couleur}">${text}</div>`,
      iconSize: [width, 28],
      iconAnchor: [half, 14],
      popupAnchor: [0, -14],
    });
  }

  // Decale legerement les marqueurs qui partagent exactement les memes
  // coordonnees, pour qu'aucun ne soit cache sous un autre.
  const duplicateCoordCounts = new Map();

  function displayLatLng(point) {
    const key = point.lat.toFixed(6) + "," + point.lon.toFixed(6);
    const n = duplicateCoordCounts.get(key) || 0;
    duplicateCoordCounts.set(key, n + 1);
    if (n === 0) return [point.lat, point.lon];

    const angle = (n - 1) * (Math.PI / 3);
    const radius = 0.00018 * (1 + Math.floor((n - 1) / 6));
    const latOffset = Math.sin(angle) * radius;
    const lonOffset = (Math.cos(angle) * radius) / Math.cos(point.lat * Math.PI / 180);
    return [point.lat + latOffset, point.lon + lonOffset];
  }

  allJours.forEach(function (jour) {
    layerGroups[jour] = L.layerGroup();
  });

  data.points.forEach(function (point) {
    const jourKey = String(point.jour);
    const group = layerGroups[jourKey];
    if (!group) return;

    const marker = L.marker(displayLatLng(point), {
      icon: createNumberIcon(point),
    });
    marker.bindTooltip(point.nom, { direction: "top", offset: [0, -16], opacity: 0.92 });
    marker.bindPopup(buildPopup(point));
    marker.pointData = point;
    group.addLayer(marker);
    markers.push(marker);
    markersByPointId.set(point.id, marker);
  });

  Object.values(layerGroups).forEach(function (group) {
    group.addTo(map);
  });

  if (markers.length > 0) {
    const bounds = L.latLngBounds(markers.map(function (m) { return m.getLatLng(); }));
    map.fitBounds(bounds.pad(0.12));
  } else {
    map.setView([52.37, 4.89], 12);
  }

  /* ---------- Etat des filtres + persistance dans l'URL ---------- */

  const filterState = {
    jours: new Set(allJours),
    segments: new Set(),
  };

  function writeStateToHash() {
    const parts = [];
    if (filterState.jours.size > 0 && filterState.jours.size < allJours.length) {
      parts.push("j=" + Array.from(filterState.jours)
        .sort(function (a, b) { return Number(a) - Number(b); })
        .join(","));
    }
    if (filterState.segments.size > 0) {
      parts.push("t=" + Array.from(filterState.segments).map(encodeURIComponent).join("!"));
    }
    const hash = parts.length ? "#" + parts.join("&") : "";
    if (window.history.replaceState) {
      window.history.replaceState(null, "", hash || window.location.pathname + window.location.search);
    } else {
      window.location.hash = hash;
    }
  }

  function readStateFromHash() {
    const hash = window.location.hash.replace(/^#/, "");
    if (!hash) return;
    hash.split("&").forEach(function (part) {
      const eq = part.indexOf("=");
      if (eq < 0) return;
      const key = part.slice(0, eq);
      const value = part.slice(eq + 1);
      if (key === "j") {
        const days = value.split(",").filter(function (d) { return allJours.indexOf(d) !== -1; });
        if (days.length) filterState.jours = new Set(days);
      } else if (key === "t") {
        value.split("!").forEach(function (raw) {
          let id = raw;
          try { id = decodeURIComponent(raw); } catch (e) { /* hash invalide, ignore */ }
          if (segmentsById.has(id)) filterState.segments.add(id);
        });
      }
    });
  }

  function pointVisible(point) {
    const jourKey = point.jour != null ? String(point.jour) : null;
    if (jourKey && filterState.jours.size > 0 && !filterState.jours.has(jourKey)) return false;
    return true;
  }

  /* ---------- Segments (trajets entre visites consecutives) ---------- */

  function buildAllSegments() {
    const segments = [];
    let colorIndex = 0;

    Object.keys(pointsByDay)
      .sort(function (a, b) { return Number(a) - Number(b); })
      .forEach(function (jourKey) {
        const points = pointsByDay[jourKey];
        for (let i = 0; i < points.length - 1; i += 1) {
          const from = points[i];
          const to = points[i + 1];
          segments.push({
            id: jourKey + ":" + from.id + ":" + to.id,
            jour: jourKey,
            from: from,
            to: to,
            label: markerLabel(from) + " → " + markerLabel(to),
            subtitle: shortName(from.nom) + " → " + shortName(to.nom),
            routeColor: ROUTE_COLORS[colorIndex % ROUTE_COLORS.length],
          });
          colorIndex += 1;
        }
      });

    return segments;
  }

  allSegments = buildAllSegments();
  allSegments.forEach(function (segment) {
    segmentsById.set(segment.id, segment);
  });

  function segmentAvailable(segment) {
    return pointVisible(segment.from) && pointVisible(segment.to);
  }

  function selectedSegments() {
    return allSegments.filter(function (segment) {
      return filterState.segments.has(segment.id) && segmentAvailable(segment);
    });
  }

  /* ---------- Calcul d'itineraires (OSRM) ---------- */

  function routeCacheKey(from, to) {
    return from.id + "->" + to.id;
  }

  function sleep(ms) {
    return new Promise(function (resolve) {
      setTimeout(resolve, ms);
    });
  }

  function decodeRouteGeometry(geometry) {
    if (!geometry || !geometry.coordinates) return [];
    return geometry.coordinates.map(function (coord) {
      return [coord[1], coord[0]];
    });
  }

  function straightLine(from, to) {
    return [
      [from.lat, from.lon],
      [to.lat, to.lon],
    ];
  }

  function airDistanceMeters(from, to) {
    const toRad = Math.PI / 180;
    const lat1 = from.lat * toRad;
    const lat2 = to.lat * toRad;
    const dLat = (to.lat - from.lat) * toRad;
    const dLon = (to.lon - from.lon) * toRad;
    const a = Math.sin(dLat / 2) * Math.sin(dLat / 2)
      + Math.cos(lat1) * Math.cos(lat2) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
    return 6371000 * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  }

  async function fetchRouteFromServer(baseUrl, from, to) {
    const url = baseUrl + "/" + from.lon + "," + from.lat + ";" + to.lon + "," + to.lat
      + "?overview=full&geometries=geojson";
    const response = await fetch(url);
    if (!response.ok) throw new Error("OSRM " + response.status);
    const payload = await response.json();
    if (payload.code !== "Ok" || !payload.routes || !payload.routes.length) {
      throw new Error(payload.message || "route introuvable");
    }
    const route = payload.routes[0];
    const latlngs = decodeRouteGeometry(route.geometry);
    if (!latlngs.length) throw new Error("geometrie vide");
    return {
      latlngs: latlngs,
      fallback: false,
      distance: route.distance,
      duration: route.duration,
      source: baseUrl.includes("openstreetmap.de") ? "piéton (OSM DE)" : "piéton (OSRM)",
    };
  }

  async function fetchRouteGeometry(from, to) {
    const cacheKey = routeCacheKey(from, to);
    if (routeCache.has(cacheKey)) {
      return routeCache.get(cacheKey);
    }

    const airDist = airDistanceMeters(from, to);
    let bestRoute = null;

    for (let i = 0; i < OSRM_SERVERS.length; i += 1) {
      try {
        const candidate = await fetchRouteFromServer(OSRM_SERVERS[i], from, to);
        if (!bestRoute || candidate.distance < bestRoute.distance) {
          bestRoute = candidate;
        }
        if (candidate.distance <= airDist * 1.8) break;
      } catch (error) {
        console.warn("Route serveur", OSRM_SERVERS[i], from.ordre_label, "->", to.ordre_label, error);
      }
    }

    if (bestRoute && bestRoute.distance > airDist * 2.5 && airDist < 1200) {
      console.warn(
        "Trajet piéton long vs distance directe",
        from.ordre_label,
        "->",
        to.ordre_label,
        Math.round(bestRoute.distance) + "m pour " + Math.round(airDist) + "m a vol d'oiseau"
      );
    }

    if (bestRoute) {
      routeCache.set(cacheKey, bestRoute);
      return bestRoute;
    }

    console.warn("Trajet approximatif", from.ordre_label, "->", to.ordre_label);
    const fallback = {
      latlngs: straightLine(from, to),
      fallback: true,
      distance: airDist,
      duration: null,
      source: "ligne droite",
    };
    routeCache.set(cacheKey, fallback);
    return fallback;
  }

  /* ---------- Affichage ---------- */

  function setTrajetsStatus(message, visible) {
    const statusEl = document.getElementById("trajets-status");
    if (!statusEl) return;
    statusEl.textContent = message || "";
    statusEl.hidden = !visible;
  }

  function fitToLatLngs(latlngs, animate) {
    if (!latlngs.length) return;
    const bounds = L.latLngBounds(latlngs).pad(0.15);
    if (animate) {
      map.flyToBounds(bounds, { duration: 0.8 });
    } else {
      map.fitBounds(bounds);
    }
  }

  function fitToVisibleMarkers(animate) {
    const visible = markers.filter(function (m) { return pointVisible(m.pointData); });
    if (visible.length > 0) {
      fitToLatLngs(visible.map(function (m) { return m.getLatLng(); }), animate !== false);
    }
  }

  function updateSegmentMeta(segment) {
    const el = segmentMetaEls.get(segment.id);
    if (!el) return;
    const route = routeCache.get(routeCacheKey(segment.from, segment.to));
    if (!route) return;
    if (route.fallback) {
      el.textContent = "≈ " + formatDistance(route.distance) + " (ligne droite)";
    } else {
      const parts = [formatDistance(route.distance)];
      if (route.duration != null) parts.push(formatDuration(route.duration));
      el.textContent = parts.join(" · ");
    }
  }

  function updateDayTotals() {
    dayTotalEls.forEach(function (el, jourKey) {
      const checked = allSegments.filter(function (segment) {
        return segment.jour === jourKey && filterState.segments.has(segment.id) && segmentAvailable(segment);
      });
      if (!checked.length) {
        el.textContent = "";
        return;
      }
      let distance = 0;
      let duration = 0;
      let hasDuration = false;
      let missing = false;
      checked.forEach(function (segment) {
        const route = routeCache.get(routeCacheKey(segment.from, segment.to));
        if (!route) {
          missing = true;
          return;
        }
        if (route.distance != null) distance += route.distance;
        if (route.duration != null) {
          duration += route.duration;
          hasDuration = true;
        }
      });
      if (missing) {
        el.textContent = "…";
        return;
      }
      const parts = [formatDistance(distance)];
      if (hasDuration) parts.push(formatDuration(duration));
      el.textContent = parts.join(" · ");
    });
  }

  async function refreshRoutes(zoomToSelection) {
    const requestId = ++routeRequestId;
    routeLayer.clearLayers();

    const segments = selectedSegments();
    if (!segments.length) {
      setTrajetsStatus("", false);
      updateDayTotals();
      return;
    }

    setTrajetsStatus("Calcul de " + segments.length + " trajet(s) à pied…", true);

    // Requetes en parallele limite, pour rester rapide sans surcharger OSRM.
    const results = new Array(segments.length);
    let nextIndex = 0;
    let done = 0;

    async function worker() {
      while (nextIndex < segments.length) {
        const i = nextIndex;
        nextIndex += 1;
        if (requestId !== routeRequestId) return;

        const segment = segments[i];
        results[i] = await fetchRouteGeometry(segment.from, segment.to);
        done += 1;

        if (requestId !== routeRequestId) return;
        setTrajetsStatus("Calcul des trajets (" + done + "/" + segments.length + ")…", true);
        updateSegmentMeta(segment);

        if (nextIndex < segments.length) {
          await sleep(ROUTE_DELAY_MS);
        }
      }
    }

    const workers = [];
    const workerCount = Math.min(ROUTE_CONCURRENCY, segments.length);
    for (let w = 0; w < workerCount; w += 1) {
      workers.push(worker());
    }
    await Promise.all(workers);

    if (requestId !== routeRequestId) return;

    const routeLatLngs = [];
    let fallbackCount = 0;

    segments.forEach(function (segment, i) {
      const route = results[i];
      if (!route) return;
      if (route.fallback) fallbackCount += 1;

      const polyline = L.polyline(route.latlngs, {
        color: segment.routeColor,
        weight: route.fallback ? 3 : 5,
        opacity: route.fallback ? 0.5 : 0.85,
        dashArray: route.fallback ? "8 8" : null,
        lineJoin: "round",
        lineCap: "round",
      });

      const distanceText = formatDistance(route.distance);
      const durationText = formatDuration(route.duration);
      const metaParts = [route.fallback ? "Trajet approximatif (ligne droite)" : "Trajet à pied (" + route.source + ")"];
      if (distanceText) metaParts.push(distanceText);
      if (durationText) metaParts.push(durationText);

      polyline.bindPopup(
        "<div class=\"popup-content\">" +
        "<span class=\"badge\">Jour " + escapeHtml(segment.from.jour) + "</span>" +
        "<p class=\"meta\"><strong>" + escapeHtml(segment.label) + "</strong></p>" +
        "<p class=\"meta\">" + escapeHtml(segment.subtitle) + "</p>" +
        "<p class=\"meta\">" + escapeHtml(metaParts.join(" · ")) + "</p>" +
        "</div>"
      );

      routeLayer.addLayer(polyline);
      route.latlngs.forEach(function (latlng) { routeLatLngs.push(latlng); });
      routeLatLngs.push([segment.from.lat, segment.from.lon], [segment.to.lat, segment.to.lon]);
    });

    updateDayTotals();

    let statusText = segments.length === 1 ? "1 trajet affiché" : segments.length + " trajets affichés";
    if (fallbackCount > 0) {
      statusText += " (dont " + fallbackCount + " approximatif" + (fallbackCount > 1 ? "s" : "") + ")";
    }
    setTrajetsStatus(statusText + ".", true);

    if (zoomToSelection) {
      fitToLatLngs(routeLatLngs, true);
    }
  }

  function refreshDayToggleLabels() {
    dayToggleButtons.forEach(function (button, jourKey) {
      const daySegments = allSegments.filter(function (segment) {
        return segment.jour === jourKey && segmentAvailable(segment);
      });
      const allChecked = daySegments.length > 0 && daySegments.every(function (segment) {
        return filterState.segments.has(segment.id);
      });
      button.textContent = allChecked ? "Tout décocher" : "Tout cocher";
    });
  }

  function syncSegmentInputs() {
    segmentInputs.forEach(function (input, segmentId) {
      const segment = segmentsById.get(segmentId);
      if (!segment) return;

      const available = segmentAvailable(segment);
      input.disabled = !available;
      input.parentElement.classList.toggle("is-disabled", !available);
      if (!available && input.checked) {
        input.checked = false;
        filterState.segments.delete(segmentId);
      }
    });
    refreshDayToggleLabels();
  }

  function refreshMarkers(zoomToMarkers) {
    markers.forEach(function (marker) {
      const point = marker.pointData;
      const group = layerGroups[String(point.jour)];
      if (pointVisible(point)) {
        if (!group.hasLayer(marker)) group.addLayer(marker);
      } else if (group.hasLayer(marker)) {
        group.removeLayer(marker);
      }
    });

    syncSegmentInputs();

    if (zoomToMarkers) {
      fitToVisibleMarkers(true);
    }

    refreshRoutes(false);
  }

  function syncDayCheckboxes() {
    dayCheckboxes.forEach(function (input, jourKey) {
      input.checked = filterState.jours.has(jourKey);
    });
  }

  function setSegmentChecked(segmentId, checked, zoomToSelection) {
    if (checked) filterState.segments.add(segmentId);
    else filterState.segments.delete(segmentId);

    const input = segmentInputs.get(segmentId);
    if (input) input.checked = checked;

    refreshDayToggleLabels();
    writeStateToHash();
    refreshRoutes(zoomToSelection);
  }

  function setDaySegments(jourKey, checked) {
    const daySegments = allSegments.filter(function (segment) {
      return segment.jour === jourKey && segmentAvailable(segment);
    });

    daySegments.forEach(function (segment) {
      if (checked) filterState.segments.add(segment.id);
      else filterState.segments.delete(segment.id);

      const input = segmentInputs.get(segment.id);
      if (input) input.checked = checked;
    });

    refreshDayToggleLabels();
    writeStateToHash();
    refreshRoutes(checked);
  }

  function clearAllSegments() {
    filterState.segments.clear();
    segmentInputs.forEach(function (input) {
      input.checked = false;
    });
    refreshDayToggleLabels();
    updateDayTotals();
    writeStateToHash();
    routeLayer.clearLayers();
    setTrajetsStatus("", false);
  }

  /* ---------- Centrage sur un point (liste des visites, navigation popup) ---------- */

  function focusPoint(pointId) {
    const marker = markersByPointId.get(pointId);
    if (!marker) return;

    const point = marker.pointData;
    const jourKey = String(point.jour);
    if (!filterState.jours.has(jourKey)) {
      filterState.jours.add(jourKey);
      syncDayCheckboxes();
      writeStateToHash();
      refreshMarkers(false);
    }

    const target = marker.getLatLng();
    if (map.getBounds().contains(target) && map.getZoom() >= 15) {
      map.panTo(target, { animate: true });
      marker.openPopup();
    } else {
      map.flyTo(target, Math.max(map.getZoom(), 16), { duration: 0.8 });
      map.once("moveend", function () {
        marker.openPopup();
      });
    }
  }

  document.addEventListener("click", function (event) {
    const btn = event.target && event.target.closest ? event.target.closest("[data-goto]") : null;
    if (!btn) return;
    focusPoint(btn.getAttribute("data-goto"));
  });

  /* ---------- Messages temporaires sur la carte ---------- */

  const mapContainerEl = document.querySelector(".map-container");
  const mapMessageEl = document.createElement("div");
  mapMessageEl.className = "map-message";
  mapMessageEl.hidden = true;
  if (mapContainerEl) mapContainerEl.appendChild(mapMessageEl);
  let mapMessageTimer = null;

  function showMapMessage(text, durationMs) {
    mapMessageEl.textContent = text;
    mapMessageEl.hidden = false;
    if (mapMessageTimer) clearTimeout(mapMessageTimer);
    mapMessageTimer = setTimeout(function () {
      mapMessageEl.hidden = true;
    }, durationMs || 4000);
  }

  /* ---------- Bouton "Ma position" ---------- */

  let locationMarker = null;
  let locationCircle = null;

  const LocateControl = L.Control.extend({
    options: { position: "topleft" },
    onAdd: function () {
      const container = L.DomUtil.create("div", "leaflet-bar leaflet-control locate-control");
      const link = L.DomUtil.create("a", "locate-button", container);
      link.href = "#";
      link.title = "Ma position";
      link.setAttribute("role", "button");
      link.setAttribute("aria-label", "Afficher ma position");
      link.innerHTML = "◎";
      L.DomEvent.on(link, "click", function (event) {
        L.DomEvent.preventDefault(event);
        L.DomEvent.stopPropagation(event);
        showMapMessage("Recherche de votre position…", 8000);
        map.locate({ enableHighAccuracy: true, timeout: 10000 });
      });
      return container;
    },
  });
  map.addControl(new LocateControl());

  map.on("locationfound", function (event) {
    const radius = Math.max(event.accuracy / 2, 15);

    if (locationMarker) map.removeLayer(locationMarker);
    if (locationCircle) map.removeLayer(locationCircle);

    locationCircle = L.circle(event.latlng, {
      radius: radius,
      color: "#2e86c1",
      weight: 1,
      fillColor: "#2e86c1",
      fillOpacity: 0.12,
    }).addTo(map);

    locationMarker = L.marker(event.latlng, {
      icon: L.divIcon({
        className: "location-marker",
        html: '<div class="location-dot" role="img" aria-label="Ma position"></div>',
        iconSize: [16, 16],
        iconAnchor: [8, 8],
      }),
    }).addTo(map);
    locationMarker.bindTooltip("Vous êtes ici", { direction: "top", offset: [0, -10] });

    mapMessageEl.hidden = true;
    map.flyTo(event.latlng, Math.max(map.getZoom(), 15), { duration: 0.8 });
  });

  map.on("locationerror", function (event) {
    console.warn("Géolocalisation impossible", event.message);
    showMapMessage("Position introuvable (autorisez la géolocalisation).");
  });

  /* ---------- Construction des filtres ---------- */

  function buildDayFilters() {
    const joursEl = document.getElementById("filter-jours");

    allJours.forEach(function (jour) {
      const id = "f-jour-" + jour;
      const row = document.createElement("div");
      row.className = "filter-day-row";

      const wrap = document.createElement("label");
      wrap.htmlFor = id;

      const input = document.createElement("input");
      input.type = "checkbox";
      input.id = id;
      input.checked = filterState.jours.has(jour);
      input.addEventListener("change", function () {
        if (input.checked) filterState.jours.add(jour);
        else filterState.jours.delete(jour);
        writeStateToHash();
        refreshMarkers(true);
      });
      dayCheckboxes.set(jour, input);

      const swatch = document.createElement("span");
      swatch.className = "day-swatch";
      swatch.style.backgroundColor = dayColor(jour);
      swatch.title = "Couleur des marqueurs du jour " + jour;

      const count = (pointsByDay[jour] || []).length;

      wrap.appendChild(input);
      wrap.appendChild(swatch);
      wrap.appendChild(document.createTextNode("Jour " + jour));

      const countEl = document.createElement("span");
      countEl.className = "filter-day-count";
      countEl.textContent = count + (count > 1 ? " visites" : " visite");
      wrap.appendChild(countEl);

      const onlyBtn = document.createElement("button");
      onlyBtn.type = "button";
      onlyBtn.className = "filter-day-only";
      onlyBtn.textContent = "seul";
      onlyBtn.title = "Afficher uniquement le jour " + jour;
      onlyBtn.addEventListener("click", function () {
        filterState.jours = new Set([jour]);
        syncDayCheckboxes();
        writeStateToHash();
        refreshMarkers(true);
      });

      row.appendChild(wrap);
      row.appendChild(onlyBtn);
      joursEl.appendChild(row);
    });
  }

  function buildVisitesList() {
    const visitesEl = document.getElementById("filter-visites");
    if (!visitesEl) return;

    Object.keys(pointsByDay)
      .sort(function (a, b) { return Number(a) - Number(b); })
      .forEach(function (jourKey) {
        const details = document.createElement("details");
        details.className = "visites-day-group";

        const summary = document.createElement("summary");
        const swatch = document.createElement("span");
        swatch.className = "day-swatch";
        swatch.style.backgroundColor = dayColor(jourKey);
        summary.appendChild(swatch);
        summary.appendChild(document.createTextNode("Jour " + jourKey));
        details.appendChild(summary);

        const list = document.createElement("div");
        list.className = "visites-items";

        pointsByDay[jourKey].forEach(function (point) {
          const btn = document.createElement("button");
          btn.type = "button";
          btn.className = "visite-item";
          btn.title = "Centrer la carte sur " + point.nom;

          const num = document.createElement("span");
          num.className = "visite-num";
          num.style.backgroundColor = point.couleur;
          num.textContent = markerLabel(point);

          const name = document.createElement("span");
          name.className = "visite-name";
          name.textContent = point.nom;

          btn.appendChild(num);
          btn.appendChild(name);
          btn.addEventListener("click", function () {
            focusPoint(point.id);
          });
          list.appendChild(btn);
        });

        details.appendChild(list);
        visitesEl.appendChild(details);
      });
  }

  function buildTrajetsFilters() {
    const trajetsEl = document.getElementById("filter-trajets");

    const segmentsByDay = {};
    allSegments.forEach(function (segment) {
      if (!segmentsByDay[segment.jour]) segmentsByDay[segment.jour] = [];
      segmentsByDay[segment.jour].push(segment);
    });

    Object.keys(segmentsByDay)
      .sort(function (a, b) { return Number(a) - Number(b); })
      .forEach(function (jourKey) {
        const details = document.createElement("details");
        details.className = "trajets-day-group";
        details.open = Number(jourKey) === Number(allJours[0]);

        const summary = document.createElement("summary");
        summary.appendChild(document.createTextNode("Jour " + jourKey));

        const total = document.createElement("span");
        total.className = "trajets-day-total";
        dayTotalEls.set(jourKey, total);
        summary.appendChild(total);

        const toggleBtn = document.createElement("button");
        toggleBtn.type = "button";
        toggleBtn.className = "trajets-day-toggle";
        toggleBtn.textContent = "Tout cocher";
        toggleBtn.addEventListener("click", function (event) {
          event.preventDefault();
          event.stopPropagation();
          const daySegments = segmentsByDay[jourKey].filter(segmentAvailable);
          const allChecked = daySegments.length > 0 && daySegments.every(function (segment) {
            return filterState.segments.has(segment.id);
          });
          setDaySegments(jourKey, !allChecked);
        });
        dayToggleButtons.set(jourKey, toggleBtn);
        summary.appendChild(toggleBtn);
        details.appendChild(summary);

        const list = document.createElement("div");
        list.className = "trajets-segments";

        segmentsByDay[jourKey].forEach(function (segment) {
          const id = "f-seg-" + segment.id.replace(/[^a-zA-Z0-9_-]/g, "_");
          const label = document.createElement("label");
          label.htmlFor = id;

          const input = document.createElement("input");
          input.type = "checkbox";
          input.id = id;
          input.checked = filterState.segments.has(segment.id);
          input.addEventListener("change", function () {
            setSegmentChecked(segment.id, input.checked, input.checked);
          });

          const swatch = document.createElement("span");
          swatch.className = "trajets-segment-color";
          swatch.style.backgroundColor = segment.routeColor;
          swatch.title = "Couleur du trajet sur la carte";

          const text = document.createElement("span");
          text.className = "trajets-segment-label";

          const strong = document.createElement("strong");
          strong.textContent = segment.label;
          const sub = document.createElement("span");
          sub.className = "trajets-segment-sub";
          sub.textContent = segment.subtitle;
          const meta = document.createElement("span");
          meta.className = "trajets-segment-meta";
          segmentMetaEls.set(segment.id, meta);

          text.appendChild(strong);
          text.appendChild(sub);
          text.appendChild(meta);

          label.appendChild(input);
          label.appendChild(swatch);
          label.appendChild(text);
          list.appendChild(label);
          segmentInputs.set(segment.id, input);
        });

        details.appendChild(list);
        trajetsEl.appendChild(details);
      });
  }

  function buildFilters() {
    buildDayFilters();
    buildVisitesList();
    buildTrajetsFilters();
  }

  /* ---------- Boutons globaux ---------- */

  document.getElementById("btn-reset").addEventListener("click", function () {
    filterState.jours = new Set(allJours);
    clearAllSegments();
    syncDayCheckboxes();
    writeStateToHash();
    refreshMarkers(true);
  });

  document.getElementById("btn-trajets-clear").addEventListener("click", function () {
    clearAllSegments();
    fitToVisibleMarkers(true);
  });

  /* ---------- Initialisation ---------- */

  readStateFromHash();
  buildFilters();
  syncSegmentInputs();

  if (filterState.jours.size < allJours.length) {
    refreshMarkers(false);
    fitToVisibleMarkers(false);
  }
  if (filterState.segments.size > 0) {
    refreshRoutes(true);
  }

  const filtersDetails = document.querySelector(".filters-details");
  if (filtersDetails) {
    if (window.matchMedia("(max-width: 768px)").matches) {
      filtersDetails.removeAttribute("open");
    }

    filtersDetails.addEventListener("toggle", function () {
      setTimeout(function () {
        map.invalidateSize();
      }, 200);
    });
  }

  const closeFiltersBtn = document.getElementById("btn-close-filters");
  if (closeFiltersBtn && filtersDetails) {
    closeFiltersBtn.addEventListener("click", function () {
      filtersDetails.removeAttribute("open");
    });
  }

  window.addEventListener("resize", function () {
    map.invalidateSize();
  });
})();
