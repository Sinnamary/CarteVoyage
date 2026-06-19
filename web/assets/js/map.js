(function () {
  "use strict";

  const data = window.VOYAGE_DATA;
  const OSRM_SERVERS = {
    foot: [
      "https://routing.openstreetmap.de/routed-foot/route/v1/foot",
      "https://router.project-osrm.org/route/v1/foot",
    ],
    car: [
      "https://routing.openstreetmap.de/routed-car/route/v1/car",
      "https://router.project-osrm.org/route/v1/driving",
    ],
  };
  // Au-delà de cette distance a vol d'oiseau, le trajet est calcule en voiture.
  const WALKING_MAX_AIR_DISTANCE_M = 5000;
  const ROUTE_DELAY_MS = 100;
  // Nombre de requetes OSRM simultanees (reste courtois envers les serveurs publics).
  const ROUTE_CONCURRENCY = 3;
  // Palette des trajets. La palette des jours (DAY_COLORS) vit dans
  // scripts/outils/excel_utils.py — champ "couleur" du JSON.
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

  /* ---------- Hébergements (base chaque soir/matin) ---------- */

  // Couleur et symbole SVG propre à chaque hébergement, identifié par la ville.
  // Le contenu SVG est dessiné dans un cercle blanc (cx=22 cy=21 r=13) au sein
  // d'une épingle 44×54 (viewBox 0 0 44 54).
  const LODGING_CONFIGS = {
    Cologne: {
      color: "#922b21",
      label: "Hôtel · Cologne",
      // Icône lit double
      svgSymbol:
        '<rect x="11" y="14" width="22" height="4" rx="2" fill="#922b21"/>' +
        '<rect x="11" y="18" width="22" height="8" rx="1" fill="#922b21"/>' +
        '<rect x="12" y="19" width="8" height="4" rx="1" fill="white" opacity="0.5"/>' +
        '<rect x="24" y="19" width="8" height="4" rx="1" fill="white" opacity="0.5"/>' +
        '<rect x="11" y="26" width="3" height="3" rx="1" fill="#922b21"/>' +
        '<rect x="30" y="26" width="3" height="3" rx="1" fill="#922b21"/>',
    },
    Amsterdam: {
      color: "#1a5276",
      label: "Hôtel · Amsterdam",
      // Icône maison néerlandaise à pignon en escalier (grachtenpand)
      svgSymbol:
        '<rect x="13" y="22" width="18" height="10" fill="#1a5276"/>' +
        '<rect x="13" y="19" width="18" height="4" fill="#1a5276"/>' +
        '<rect x="15" y="16" width="14" height="4" fill="#1a5276"/>' +
        '<rect x="17" y="13" width="10" height="4" fill="#1a5276"/>' +
        '<rect x="19" y="10" width="6" height="4" fill="#1a5276"/>' +
        '<rect x="19" y="26" width="6" height="6" rx="1" fill="white" opacity="0.75"/>',
    },
    Ennevelin: {
      color: "#145a32",
      label: "Maison · Ennevelin",
      // Icône maison de campagne avec cheminée
      svgSymbol:
        '<rect x="13" y="22" width="18" height="10" fill="#145a32"/>' +
        '<path d="M22,10 L32,22 L12,22 Z" fill="#145a32"/>' +
        '<rect x="25" y="10" width="4" height="8" fill="#145a32"/>' +
        '<rect x="14" y="23" width="5" height="5" rx="0.5" fill="white" opacity="0.7"/>' +
        '<rect x="19" y="26" width="6" height="6" rx="1" fill="white" opacity="0.75"/>',
    },
  };

  function isLodgingPoint(point) {
    const action = point.popup && point.popup.action;
    return action != null && String(action).trim().toLowerCase() === "hébergement";
  }

  function lodgingConfig(point) {
    return LODGING_CONFIGS[point.ville] || {
      color: "#6c3483",
      label: "Hébergement",
      svgSymbol:
        '<path d="M22,12 L30,22 L14,22 Z" fill="#6c3483"/>' +
        '<rect x="15" y="22" width="14" height="9" fill="#6c3483"/>' +
        '<rect x="19" y="25" width="6" height="6" rx="1" fill="white" opacity="0.75"/>',
    };
  }

  function createLodgingIcon(point) {
    const cfg = lodgingConfig(point);
    const ariaLabel = escapeHtml("Hébergement — " + point.nom);
    const svg =
      `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 44 54" width="44" height="54">` +
      `<path d="M22,3 C12,3 4,11 4,21 C4,32 22,51 22,51 C22,51 40,32 40,21 C40,11 32,3 22,3 Z"` +
      ` fill="${cfg.color}" stroke="white" stroke-width="2.5"/>` +
      `<circle cx="22" cy="21" r="13" fill="white" opacity="0.95"/>` +
      cfg.svgSymbol +
      `</svg>`;
    return L.divIcon({
      className: "lodging-marker",
      html: `<div role="img" aria-label="${ariaLabel}">${svg}</div>`,
      iconSize: [44, 54],
      iconAnchor: [22, 51],
      popupAnchor: [0, -54],
    });
  }

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
  const OSM_CACHE_STORAGE_KEY = "cartevoyage-osm-routes-v1";
  const osmPersistentCache = new Map();
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

  /* ---------- Registre fusionné des hébergements ---------- */
  // Regroupe les points Hébergement par (jour × nom) → un seul marqueur par base.
  // first  = départ le matin  (visite à l'index le plus bas)
  // last   = arrivée le soir  (visite à l'index le plus haut)
  // Utilise la POSITION (visite), pas les horaires, pour éviter les erreurs de saisie.
  const lodgingRegistry = new Map(); // key "jour:nom"

  data.points.forEach(function (point) {
    if (!isLodgingPoint(point)) return;
    const key = point.jour + ":" + point.nom;
    if (!lodgingRegistry.has(key)) {
      lodgingRegistry.set(key, { first: null, last: null, ids: [], marker: null });
    }
    const entry = lodgingRegistry.get(key);
    entry.ids.push(point.id);
    if (!entry.first || point.visite < entry.first.visite) entry.first = point;
    if (!entry.last || point.visite > entry.last.visite) entry.last = point;
  });

  /* ---------- Popups et marqueurs ---------- */

  function buildMergedLodgingPopup(entry, point) {
    const cfg = lodgingConfig(point);
    const first = entry.first;   // départ matin
    const last = entry.last;     // arrivée soir

    const svgIcon =
      `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 44 54" width="28" height="34" class="lodging-popup-icon">` +
      `<path d="M22,3 C12,3 4,11 4,21 C4,32 22,51 22,51 C22,51 40,32 40,21 C40,11 32,3 22,3 Z"` +
      ` fill="${cfg.color}" stroke="white" stroke-width="2.5"/>` +
      `<circle cx="22" cy="21" r="13" fill="white" opacity="0.95"/>` +
      cfg.svgSymbol +
      `</svg>`;

    const parts = [
      `<div class="popup-content">`,
      `<span class="badge">Jour ${escapeHtml(point.jour)} · Base nuitée</span>`,
      `<div class="lodging-popup-header">`,
      svgIcon,
      `<h3 class="lodging-popup-title">${escapeHtml(point.nom)}</h3>`,
      `</div>`,
      `<p class="meta lodging-base-label">${escapeHtml(cfg.label)}</p>`,
    ];

    // Heure de départ (matin) et arrivée (soir)
    if (first && last && first.id !== last.id) {
      // Deux entrées : matin + soir
      const departHeure = (first.popup && first.popup.fermeture) || "";
      const arriveeHeure = (last.popup && last.popup.ouverture) || "";
      if (departHeure) {
        parts.push(`<p class="meta">&#9728; Départ : ${escapeHtml(departHeure)}</p>`);
      }
      if (arriveeHeure) {
        parts.push(`<p class="meta">&#9790; Arrivée : ${escapeHtml(arriveeHeure)}</p>`);
      }
    } else {
      // Une seule entrée
      const p = point.popup || {};
      if (p.ouverture || p.fermeture) {
        parts.push(`<p class="meta">Horaires : ${escapeHtml(p.ouverture || "?")} – ${escapeHtml(p.fermeture || "?")}</p>`);
      }
    }

    // Lien réservation
    const lien = (last && last.lien) || (first && first.lien) || null;
    if (lien) {
      parts.push(`<p class="meta"><a href="${escapeHtml(lien)}" target="_blank" rel="noopener">Réservation</a></p>`);
    }

    // Navigation vers la première / dernière visite du jour
    const neighborsFirst = first ? (neighborsById.get(first.id) || {}) : {};
    const neighborsLast = last ? (neighborsById.get(last.id) || {}) : {};
    const firstVisit = neighborsFirst.next && !isLodgingPoint(neighborsFirst.next) ? neighborsFirst.next : null;
    const lastVisit = neighborsLast.prev && !isLodgingPoint(neighborsLast.prev) ? neighborsLast.prev : null;
    if (firstVisit || lastVisit) {
      parts.push(`<div class="popup-nav">`);
      if (firstVisit) {
        parts.push(
          `<button type="button" class="popup-nav-btn" data-goto="${escapeHtml(firstVisit.id)}" ` +
          `title="${escapeHtml(firstVisit.nom)}">1ère visite &rarr;</button>`
        );
      } else {
        parts.push(`<span></span>`);
      }
      if (lastVisit) {
        parts.push(
          `<button type="button" class="popup-nav-btn" data-goto="${escapeHtml(lastVisit.id)}" ` +
          `title="${escapeHtml(lastVisit.nom)}">&larr; Dernière visite</button>`
        );
      }
      parts.push(`</div>`);
    }

    parts.push(`</div>`);
    return parts.join("");
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
    if (isLodgingPoint(point)) {
      return createLodgingIcon(point);
    }
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

    if (isLodgingPoint(point)) {
      const lodgingKey = point.jour + ":" + point.nom;
      const entry = lodgingRegistry.get(lodgingKey);
      if (!entry) return;

      if (entry.marker) {
        // Marqueur déjà créé pour cet hébergement ce jour : on enregistre juste l'id.
        markersByPointId.set(point.id, entry.marker);
        return;
      }

      // Créer le marqueur fusionné (coordonnées exactes, sans décalage).
      const displayPoint = entry.last || entry.first || point;
      const marker = L.marker([displayPoint.lat, displayPoint.lon], {
        icon: createLodgingIcon(displayPoint),
        zIndexOffset: 500, // au-dessus des marqueurs normaux
      });
      marker.bindTooltip(displayPoint.nom, { direction: "top", offset: [0, -24], opacity: 0.92 });
      marker.bindPopup(buildMergedLodgingPopup(entry, displayPoint));
      marker.pointData = displayPoint;

      group.addLayer(marker);
      markers.push(marker);
      entry.marker = marker;
      entry.ids.forEach(function (id) { markersByPointId.set(id, marker); });
      return;
    }

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
    excludeCarRoutes: false,
    routeCalculation: "osm",
  };

  function isCarSegment(segment) {
    return segment.mode === "car";
  }

  function displayedSegments(segments) {
    const list = segments || selectedSegments();
    if (!filterState.excludeCarRoutes) return list;
    return list.filter(function (segment) { return !isCarSegment(segment); });
  }

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
    if (!filterState.excludeCarRoutes) {
      parts.push("c=0");
    }
    if (filterState.routeCalculation === "air") {
      parts.push("r=air");
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
      } else if (key === "c" && value === "0") {
        filterState.excludeCarRoutes = false;
      } else if (key === "r" && value === "air") {
        filterState.routeCalculation = "air";
      }
    });
  }

  function pointHiddenByCarFilter(point) {
    if (!filterState.excludeCarRoutes) return false;
    if (isTransportPoint(point)) return true;
    const segments = allSegments.filter(function (segment) {
      return segment.jour === String(point.jour)
        && (segment.from.id === point.id || segment.to.id === point.id);
    });
    if (!segments.length) return false;
    return segments.every(function (segment) { return isCarSegment(segment); });
  }

  function pointVisible(point) {
    const jourKey = point.jour != null ? String(point.jour) : null;
    if (jourKey && filterState.jours.size > 0 && !filterState.jours.has(jourKey)) return false;
    if (pointHiddenByCarFilter(point)) return false;
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
          const mode = segmentTravelMode(from, to);
          segments.push({
            id: jourKey + ":" + from.id + ":" + to.id,
            jour: jourKey,
            from: from,
            to: to,
            mode: mode,
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

  function normalizeVille(ville) {
    if (ville == null || ville === "") return "";
    return String(ville).trim().toLowerCase();
  }

  function isTransportPoint(point) {
    const action = point.popup && point.popup.action;
    return action != null && String(action).trim().toLowerCase() === "transport";
  }

  function segmentTravelMode(from, to) {
    if (isTransportPoint(from) || isTransportPoint(to)) return "car";

    const villeFrom = normalizeVille(from.ville);
    const villeTo = normalizeVille(to.ville);
    if (villeFrom && villeTo && villeFrom !== villeTo) return "car";

    if (airDistanceMeters(from, to) > WALKING_MAX_AIR_DISTANCE_M) return "car";

    return "foot";
  }

  function travelModeLabel(mode) {
    return mode === "car" ? "en voiture" : "à pied";
  }

  function routeCacheKey(from, to, mode) {
    return filterState.routeCalculation + ":" + mode + ":" + from.id + "->" + to.id;
  }

  function osmCacheKey(from, to, mode) {
    return mode + ":"
      + from.lat.toFixed(5) + "," + from.lon.toFixed(5) + "->"
      + to.lat.toFixed(5) + "," + to.lon.toFixed(5);
  }

  function loadOsmCacheFromStorage() {
    try {
      const raw = localStorage.getItem(OSM_CACHE_STORAGE_KEY);
      if (!raw) return;
      const data = JSON.parse(raw);
      if (typeof data !== "object" || !data) return;
      Object.keys(data).forEach(function (key) {
        const route = data[key];
        if (route && Array.isArray(route.latlngs)) {
          osmPersistentCache.set(key, route);
        }
      });
    } catch (error) {
      console.warn("Cache OSM local illisible", error);
    }
  }

  function persistOsmRoute(from, to, mode, route) {
    if (!route || route.fallback || route.airMode) return;
    const key = osmCacheKey(from, to, mode);
    osmPersistentCache.set(key, route);
    try {
      const data = {};
      osmPersistentCache.forEach(function (value, cacheKey) {
        data[cacheKey] = value;
      });
      localStorage.setItem(OSM_CACHE_STORAGE_KEY, JSON.stringify(data));
    } catch (error) {
      console.warn("Cache OSM local non enregistre", error);
    }
  }

  function getOsmPersistentRoute(from, to, mode) {
    return osmPersistentCache.get(osmCacheKey(from, to, mode)) || null;
  }

  function isRouteCached(from, to, mode) {
    return routeCache.has(routeCacheKey(from, to, mode));
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

  function buildAirRoute(from, to, mode) {
    const airDist = airDistanceMeters(from, to);
    return {
      latlngs: straightLine(from, to),
      fallback: false,
      airMode: true,
      distance: airDist,
      duration: mode === "car" ? Math.round(airDist / 22) : Math.round(airDist / 1.4),
      mode: mode,
      source: "vol d'oiseau",
    };
  }

  async function fetchRouteFromServer(baseUrl, from, to, mode) {
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
    const isCar = mode === "car";
    const osmDe = baseUrl.includes("openstreetmap.de");
    return {
      latlngs: latlngs,
      fallback: false,
      distance: route.distance,
      duration: route.duration,
      mode: mode,
      source: isCar
        ? (osmDe ? "voiture (OSM DE)" : "voiture (OSRM)")
        : (osmDe ? "piéton (OSM DE)" : "piéton (OSRM)"),
    };
  }

  async function fetchRouteGeometry(from, to, mode) {
    const cacheKey = routeCacheKey(from, to, mode);
    if (routeCache.has(cacheKey)) {
      return routeCache.get(cacheKey);
    }

    if (filterState.routeCalculation === "air") {
      const route = buildAirRoute(from, to, mode);
      routeCache.set(cacheKey, route);
      return route;
    }

    const storedRoute = getOsmPersistentRoute(from, to, mode);
    if (storedRoute) {
      routeCache.set(cacheKey, storedRoute);
      return storedRoute;
    }

    const airDist = airDistanceMeters(from, to);
    const servers = OSRM_SERVERS[mode] || OSRM_SERVERS.foot;
    let bestRoute = null;

    for (let i = 0; i < servers.length; i += 1) {
      try {
        const candidate = await fetchRouteFromServer(servers[i], from, to, mode);
        if (!bestRoute || candidate.distance < bestRoute.distance) {
          bestRoute = candidate;
        }
        if (candidate.distance <= airDist * 1.8) break;
      } catch (error) {
        console.warn("Route serveur", servers[i], from.ordre_label, "->", to.ordre_label, error);
      }
    }

    if (mode === "foot" && bestRoute && bestRoute.distance > airDist * 2.5 && airDist < 1200) {
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
      persistOsmRoute(from, to, mode, bestRoute);
      return bestRoute;
    }

    console.warn("Trajet approximatif", from.ordre_label, "->", to.ordre_label);
    const fallback = {
      latlngs: straightLine(from, to),
      fallback: true,
      distance: airDist,
      duration: mode === "car" ? Math.round(airDist / 22) : null,
      mode: mode,
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
    const route = routeCache.get(routeCacheKey(segment.from, segment.to, segment.mode));
    if (!route) return;
    if (route.airMode) {
      const parts = ["≈ " + formatDistance(route.distance), "vol d'oiseau"];
      if (route.duration != null) parts.splice(1, 0, formatDuration(route.duration));
      el.textContent = parts.join(" · ");
    } else if (route.fallback) {
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
        return segment.jour === jourKey
          && filterState.segments.has(segment.id)
          && segmentAvailable(segment);
      });
      const visible = displayedSegments(checked);
      if (!visible.length) {
        el.textContent = checked.length ? "—" : "";
        return;
      }
      let distance = 0;
      let duration = 0;
      let hasDuration = false;
      let missing = false;
      visible.forEach(function (segment) {
        const route = routeCache.get(routeCacheKey(segment.from, segment.to, segment.mode));
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
    const drawnSegments = displayedSegments(segments);
    if (!segments.length) {
      setTrajetsStatus("", false);
      updateDayTotals();
      return;
    }

    const uncachedSegments = segments.filter(function (segment) {
      return !isRouteCached(segment.from, segment.to, segment.mode);
    });

    let results;
    if (uncachedSegments.length === 0) {
      results = segments.map(function (segment) {
        return routeCache.get(routeCacheKey(segment.from, segment.to, segment.mode));
      });
      segments.forEach(updateSegmentMeta);
    } else {
      setTrajetsStatus("Calcul de " + uncachedSegments.length + " trajet(s)…", true);

      results = new Array(segments.length);
      let nextIndex = 0;
      let fetchedCount = 0;

      async function worker() {
        while (nextIndex < segments.length) {
          const i = nextIndex;
          nextIndex += 1;
          if (requestId !== routeRequestId) return;

          const segment = segments[i];
          const wasCached = isRouteCached(segment.from, segment.to, segment.mode);
          results[i] = await fetchRouteGeometry(segment.from, segment.to, segment.mode);

          if (requestId !== routeRequestId) return;
          if (!wasCached) {
            fetchedCount += 1;
            if (filterState.routeCalculation === "osm") {
              setTrajetsStatus(
                "Calcul des trajets (" + fetchedCount + "/" + uncachedSegments.length + ")…",
                true
              );
            }
          }
          updateSegmentMeta(segment);

          if (!wasCached && filterState.routeCalculation === "osm") {
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
    }

    if (requestId !== routeRequestId) return;

    const routeLatLngs = [];
    let fallbackCount = 0;
    const drawnIds = new Set(drawnSegments.map(function (segment) { return segment.id; }));

    segments.forEach(function (segment, i) {
      const route = results[i];
      if (!route) return;
      if (!drawnIds.has(segment.id)) return;
      if (route.fallback) fallbackCount += 1;

      const isCar = segment.mode === "car";
      const isAir = !!route.airMode;
      const isApprox = route.fallback || isAir;
      const polyline = L.polyline(route.latlngs, {
        color: segment.routeColor,
        weight: isApprox ? 3 : (isCar ? 4 : 5),
        opacity: isApprox ? 0.55 : 0.85,
        dashArray: isApprox ? "8 8" : (isCar ? "10 6" : null),
        lineJoin: "round",
        lineCap: "round",
      });

      const distanceText = formatDistance(route.distance);
      const durationText = formatDuration(route.duration);
      const modeText = travelModeLabel(segment.mode);
      const metaParts = [isAir
        ? "Trajet " + modeText + " (vol d'oiseau)"
        : route.fallback
          ? "Trajet approximatif (ligne droite, " + modeText + ")"
          : "Trajet " + modeText + " (" + route.source + ")"];
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

    const hiddenCarCount = segments.length - drawnSegments.length;
    let statusText;
    if (!drawnSegments.length) {
      statusText = hiddenCarCount > 0
        ? "Trajets voiture masqués — carte centrée sur les activités"
        : "Aucun trajet affiché";
    } else {
      statusText = drawnSegments.length === 1
        ? "1 trajet affiché"
        : drawnSegments.length + " trajets affichés";
      if (hiddenCarCount > 0) {
        statusText += " (" + hiddenCarCount + " trajet(s) voiture masqué(s))";
      }
    }
    if (fallbackCount > 0) {
      statusText += " (dont " + fallbackCount + " approximatif" + (fallbackCount > 1 ? "s" : "") + ")";
    }
    if (filterState.routeCalculation === "air") {
      statusText += " · vol d'oiseau";
    }
    setTrajetsStatus(statusText + ".", true);

    if (zoomToSelection) {
      if (routeLatLngs.length > 0) {
        fitToLatLngs(routeLatLngs, true);
      } else {
        fitToVisibleMarkers(true);
      }
    }
  }

  function refreshDayToggleLabels() {
    dayToggleButtons.forEach(function (button, jourKey) {
      const daySegments = allSegments.filter(function (segment) {
        return segment.jour === jourKey && segmentAvailable(segment);
      });
      const targetSegments = filterState.excludeCarRoutes
        ? daySegments.filter(function (segment) { return !isCarSegment(segment); })
        : daySegments;
      const allChecked = targetSegments.length > 0 && targetSegments.every(function (segment) {
        return filterState.segments.has(segment.id);
      });
      button.textContent = allChecked ? "Tout décocher" : "Tout cocher";
    });
  }

  function syncVisitesCarFilter() {
    document.querySelectorAll(".visite-item[data-point-id]").forEach(function (btn) {
      const marker = markersByPointId.get(btn.getAttribute("data-point-id"));
      if (!marker) return;
      const point = marker.pointData;
      const jourKey = String(point.jour);
      const jourHidden = filterState.jours.size > 0 && !filterState.jours.has(jourKey);
      const carHidden = pointHiddenByCarFilter(point);
      btn.classList.toggle("is-car-hidden", !jourHidden && carHidden);
      btn.disabled = carHidden;
    });
  }

  function syncCarSegmentStyles() {
    segmentInputs.forEach(function (input, segmentId) {
      const segment = segmentsById.get(segmentId);
      if (!segment || !input.parentElement) return;
      const hidden = filterState.excludeCarRoutes
        && isCarSegment(segment)
        && filterState.segments.has(segmentId);
      input.parentElement.classList.toggle("is-car-hidden", hidden);
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
    syncCarSegmentStyles();
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
    syncVisitesCarFilter();

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
    syncCarSegmentStyles();
  }

  function setDaySegments(jourKey, checked) {
    const daySegments = allSegments.filter(function (segment) {
      return segment.jour === jourKey && segmentAvailable(segment);
    });

    daySegments.forEach(function (segment) {
      if (checked && filterState.excludeCarRoutes && isCarSegment(segment)) return;
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

      const dayPoints = pointsByDay[jour] || [];
      const count = dayPoints.length;

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

        const shownLodgingKeys = new Set();

        pointsByDay[jourKey].forEach(function (point) {
          const isLodging = isLodgingPoint(point);

          if (isLodging) {
            // Ne montrer qu'un seul bouton par hébergement fusionné.
            const lodgingKey = point.jour + ":" + point.nom;
            if (shownLodgingKeys.has(lodgingKey)) return;
            shownLodgingKeys.add(lodgingKey);
          }

          const btn = document.createElement("button");
          btn.type = "button";
          btn.dataset.pointId = point.id;
          btn.title = "Centrer la carte sur " + point.nom;
          btn.className = isLodging ? "visite-item visite-item--lodging" : "visite-item";

          const num = document.createElement("span");
          if (isLodging) {
            const cfg = lodgingConfig(point);
            num.className = "visite-num visite-num--lodging";
            num.innerHTML =
              `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 44 54" width="16" height="20">` +
              `<path d="M22,3 C12,3 4,11 4,21 C4,32 22,51 22,51 C22,51 40,32 40,21 C40,11 32,3 22,3 Z"` +
              ` fill="${cfg.color}" stroke="white" stroke-width="3"/>` +
              `<circle cx="22" cy="21" r="13" fill="white" opacity="0.9"/>` +
              cfg.svgSymbol +
              `</svg>`;
          } else {
            num.className = "visite-num";
            num.style.backgroundColor = point.couleur;
            num.textContent = markerLabel(point);
          }

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
          sub.textContent = segment.subtitle + " · " + travelModeLabel(segment.mode);
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

  const excludeCarToggle = document.getElementById("toggle-exclude-car");
  const routeModeOsmBtn = document.getElementById("route-mode-osm");
  const routeModeAirBtn = document.getElementById("route-mode-air");

  function syncRouteModeButtons() {
    const isAir = filterState.routeCalculation === "air";
    if (routeModeOsmBtn) {
      routeModeOsmBtn.classList.toggle("is-active", !isAir);
      routeModeOsmBtn.setAttribute("aria-pressed", String(!isAir));
    }
    if (routeModeAirBtn) {
      routeModeAirBtn.classList.toggle("is-active", isAir);
      routeModeAirBtn.setAttribute("aria-pressed", String(isAir));
    }
  }

  function setRouteCalculation(mode) {
    if (mode !== "osm" && mode !== "air") return;
    if (filterState.routeCalculation === mode) return;
    filterState.routeCalculation = mode;
    syncRouteModeButtons();
    writeStateToHash();
    if (filterState.segments.size > 0) {
      refreshRoutes(false);
    } else {
      updateDayTotals();
    }
  }

  if (routeModeOsmBtn) {
    routeModeOsmBtn.addEventListener("click", function () {
      setRouteCalculation("osm");
    });
  }
  if (routeModeAirBtn) {
    routeModeAirBtn.addEventListener("click", function () {
      setRouteCalculation("air");
    });
  }

  document.getElementById("btn-reset").addEventListener("click", function () {
    filterState.jours = new Set(allJours);
    filterState.excludeCarRoutes = false;
    filterState.routeCalculation = "osm";
    if (excludeCarToggle) excludeCarToggle.checked = false;
    syncRouteModeButtons();
    clearAllSegments();
    syncDayCheckboxes();
    writeStateToHash();
    refreshMarkers(true);
  });

  document.getElementById("btn-trajets-clear").addEventListener("click", function () {
    clearAllSegments();
    fitToVisibleMarkers(true);
  });

  if (excludeCarToggle) {
    excludeCarToggle.checked = filterState.excludeCarRoutes;
    excludeCarToggle.addEventListener("change", function () {
      filterState.excludeCarRoutes = excludeCarToggle.checked;
      syncCarSegmentStyles();
      syncVisitesCarFilter();
      writeStateToHash();
      refreshMarkers(true);
    });
  }

  /* ---------- Initialisation ---------- */

  loadOsmCacheFromStorage();
  readStateFromHash();
  syncRouteModeButtons();
  buildFilters();
  syncSegmentInputs();
  syncCarSegmentStyles();
  syncVisitesCarFilter();
  refreshMarkers(false);

  if (filterState.jours.size < allJours.length) {
    fitToVisibleMarkers(false);
  } else if (filterState.excludeCarRoutes) {
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
