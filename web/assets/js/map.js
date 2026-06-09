(function () {
  "use strict";

  const data = window.VOYAGE_DATA;
  const OSRM_SERVERS = [
    "https://routing.openstreetmap.de/routed-foot/route/v1/foot",
    "https://router.project-osrm.org/route/v1/foot",
  ];
  const ROUTE_DELAY_MS = 120;
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

  if (!data || !data.points) {
    console.error("VOYAGE_DATA manquant");
    return;
  }

  const map = L.map("map", { zoomControl: true });
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
    maxZoom: 19,
  }).addTo(map);

  const layerGroups = {};
  const routeLayer = L.layerGroup().addTo(map);
  const markers = [];
  const routeCache = new Map();
  const segmentInputs = new Map();
  const segmentsById = new Map();
  const dayToggleButtons = new Map();
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
    parts.push(`</div>`);
    return parts.join("");
  }

  function createNumberIcon(label, color) {
    const text = escapeHtml(label);
    const width = Math.max(28, String(label).length * 9 + 14);
    const half = Math.round(width / 2);
    return L.divIcon({
      className: "numbered-marker",
      html: `<div class="marker-number" style="background:${color}">${text}</div>`,
      iconSize: [width, 28],
      iconAnchor: [half, 14],
      popupAnchor: [0, -14],
    });
  }

  const allJours = (data.jours || [])
    .map(function (jour) { return String(jour); })
    .sort(function (a, b) { return Number(a) - Number(b); });

  allJours.forEach(function (jour) {
    layerGroups[jour] = L.layerGroup();
  });

  data.points.forEach(function (point) {
    const jourKey = String(point.jour);
    const group = layerGroups[jourKey];
    if (!group) return;

    const marker = L.marker([point.lat, point.lon], {
      icon: createNumberIcon(markerLabel(point), point.couleur),
    });
    marker.bindPopup(buildPopup(point));
    marker.pointData = point;
    group.addLayer(marker);
    markers.push(marker);
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

  const filterState = {
    jours: new Set(allJours),
    segments: new Set(),
  };

  function pointVisible(point) {
    const jourKey = point.jour != null ? String(point.jour) : null;
    if (jourKey && filterState.jours.size > 0 && !filterState.jours.has(jourKey)) return false;
    return true;
  }

  function buildAllSegments() {
    const byDay = {};
    data.points.forEach(function (point) {
      const jourKey = String(point.jour);
      if (!byDay[jourKey]) byDay[jourKey] = [];
      byDay[jourKey].push(point);
    });

    const segments = [];
    let colorIndex = 0;

    Object.keys(byDay)
      .sort(function (a, b) { return Number(a) - Number(b); })
      .forEach(function (jourKey) {
        const points = byDay[jourKey].sort(function (a, b) {
          return a.visite - b.visite || String(a.id).localeCompare(String(b.id));
        });

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

  function setTrajetsStatus(message, visible) {
    const statusEl = document.getElementById("trajets-status");
    if (!statusEl) return;
    statusEl.textContent = message || "";
    statusEl.hidden = !visible;
  }

  function fitToLatLngs(latlngs) {
    if (!latlngs.length) return;
    const bounds = L.latLngBounds(latlngs);
    map.fitBounds(bounds.pad(0.15));
  }

  function fitToVisibleMarkers() {
    const visible = markers.filter(function (m) { return pointVisible(m.pointData); });
    if (visible.length > 0) {
      fitToLatLngs(visible.map(function (m) { return m.getLatLng(); }));
    }
  }

  async function refreshRoutes(zoomToSelection) {
    const requestId = ++routeRequestId;
    routeLayer.clearLayers();

    const segments = selectedSegments();
    if (!segments.length) {
      setTrajetsStatus("", false);
      return;
    }

    setTrajetsStatus("Calcul de " + segments.length + " trajet(s) a pied…", true);

    const routeLatLngs = [];
    let done = 0;

    for (let i = 0; i < segments.length; i += 1) {
      if (requestId !== routeRequestId) return;

      const segment = segments[i];
      const route = await fetchRouteGeometry(segment.from, segment.to);
      done += 1;
      setTrajetsStatus("Calcul des trajets (" + done + "/" + segments.length + ")…", true);

      if (requestId !== routeRequestId) return;

      const polyline = L.polyline(route.latlngs, {
        color: segment.routeColor,
        weight: route.fallback ? 3 : 5,
        opacity: route.fallback ? 0.5 : 0.85,
        dashArray: route.fallback ? "8 8" : null,
        lineJoin: "round",
        lineCap: "round",
      });

      const distanceText = route.distance != null
        ? Math.round(route.distance) + " m"
        : "";
      const durationText = route.duration != null
        ? Math.round(route.duration / 60) + " min"
        : "";
      const metaParts = [route.fallback ? "Trajet approximatif (ligne droite)" : "Trajet a pied (" + route.source + ")"];
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

      if (i < segments.length - 1) {
        await sleep(ROUTE_DELAY_MS);
      }
    }

    if (requestId !== routeRequestId) return;

    setTrajetsStatus(
      segments.length === 1
        ? "1 trajet affiche."
        : segments.length + " trajets affiches.",
      true
    );

    if (zoomToSelection) {
      fitToLatLngs(routeLatLngs);
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
      button.textContent = allChecked ? "Tout decocher" : "Tout cocher";
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
      fitToVisibleMarkers();
    }

    const hasRoutes = filterState.segments.size > 0;
    refreshRoutes(hasRoutes);
  }

  function setSegmentChecked(segmentId, checked, zoomToSelection) {
    if (checked) filterState.segments.add(segmentId);
    else filterState.segments.delete(segmentId);

    const input = segmentInputs.get(segmentId);
    if (input) input.checked = checked;

    refreshDayToggleLabels();
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
    refreshRoutes(checked && daySegments.length === 1);
  }

  function clearAllSegments() {
    filterState.segments.clear();
    segmentInputs.forEach(function (input) {
      input.checked = false;
    });
    refreshDayToggleLabels();
    routeLayer.clearLayers();
    setTrajetsStatus("", false);
  }

  function buildFilters() {
    const joursEl = document.getElementById("filter-jours");
    const trajetsEl = document.getElementById("filter-trajets");

    allJours.forEach(function (jour) {
      const id = "f-jour-" + jour;
      const wrap = document.createElement("label");
      wrap.htmlFor = id;
      const input = document.createElement("input");
      input.type = "checkbox";
      input.id = id;
      input.checked = filterState.jours.has(jour);
      input.addEventListener("change", function () {
        if (input.checked) filterState.jours.add(jour);
        else filterState.jours.delete(jour);
        refreshMarkers(true);
      });
      wrap.appendChild(input);
      wrap.appendChild(document.createTextNode("Jour " + jour));
      joursEl.appendChild(wrap);
    });

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
          text.innerHTML =
            "<strong>" + escapeHtml(segment.label) + "</strong>" +
            "<span class=\"trajets-segment-sub\">" + escapeHtml(segment.subtitle) + "</span>";

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

  document.getElementById("btn-reset").addEventListener("click", function () {
    filterState.jours = new Set(allJours);
    clearAllSegments();
    document.querySelectorAll("#filter-jours input[type=checkbox]").forEach(function (input) {
      input.checked = true;
    });
    refreshMarkers(true);
  });

  document.getElementById("btn-trajets-clear").addEventListener("click", function () {
    clearAllSegments();
    fitToVisibleMarkers();
  });

  buildFilters();
  syncSegmentInputs();

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

  window.addEventListener("resize", function () {
    map.invalidateSize();
  });
})();
