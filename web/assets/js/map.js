(function () {
  "use strict";

  const data = window.VOYAGE_DATA;
  const pageFilter = window.PAGE_FILTER || { mode: "all" };

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
  const markers = [];

  function escapeHtml(text) {
    if (text == null) return "";
    const div = document.createElement("div");
    div.textContent = String(text);
    return div.innerHTML;
  }

  function buildPopup(point) {
    const p = point.popup || {};
    const parts = [
      `<div class="popup-content">`,
      `<span class="badge">${escapeHtml(point.onglet)}${point.jour ? " — J" + point.jour : ""}</span>`,
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

  function createNumberIcon(number, color) {
    return L.divIcon({
      className: "numbered-marker",
      html: `<div class="marker-number" style="background:${color}">${number}</div>`,
      iconSize: [28, 28],
      iconAnchor: [14, 14],
      popupAnchor: [0, -14],
    });
  }

  data.onglets.forEach(function (onglet) {
    layerGroups[onglet] = L.layerGroup();
  });

  data.points.forEach(function (point) {
    const group = layerGroups[point.onglet];
    if (!group) return;

    const marker = L.marker([point.lat, point.lon], {
      icon: createNumberIcon(point.ordre, point.couleur),
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
    villes: new Set(data.villes),
    jours: new Set(Object.keys(data.jours || {})),
    onglets: new Set(data.onglets),
  };

  function applyPageFilter() {
    if (pageFilter.mode === "ville" && pageFilter.ville) {
      filterState.villes = new Set([pageFilter.ville]);
    }
  }

  function pointVisible(point) {
    const jourKey = point.jour != null ? String(point.jour) : null;
    if (!filterState.villes.has(point.ville)) return false;
    if (jourKey && filterState.jours.size > 0 && !filterState.jours.has(jourKey)) return false;
    if (!filterState.onglets.has(point.onglet)) return false;
    return true;
  }

  function refreshMarkers() {
    markers.forEach(function (marker) {
      const point = marker.pointData;
      const group = layerGroups[point.onglet];
      if (pointVisible(point)) {
        if (!group.hasLayer(marker)) group.addLayer(marker);
      } else if (group.hasLayer(marker)) {
        group.removeLayer(marker);
      }
    });

    const visible = markers.filter(function (m) { return pointVisible(m.pointData); });
    if (visible.length > 0) {
      const bounds = L.latLngBounds(visible.map(function (m) { return m.getLatLng(); }));
      map.fitBounds(bounds.pad(0.12));
    }
  }

  function makeCheckbox(container, label, value, checked, onChange) {
    const id = "f-" + label.replace(/\s+/g, "-").toLowerCase() + "-" + value;
    const wrap = document.createElement("label");
    wrap.htmlFor = id;
    const input = document.createElement("input");
    input.type = "checkbox";
    input.id = id;
    input.checked = checked;
    input.addEventListener("change", function () {
      onChange(input.checked);
      refreshMarkers();
    });
    wrap.appendChild(input);
    wrap.appendChild(document.createTextNode(label));
    container.appendChild(wrap);
  }

  function buildFilters() {
    const villesEl = document.getElementById("filter-villes");
    const joursEl = document.getElementById("filter-jours");
    const ongletsEl = document.getElementById("filter-onglets");

    data.villes.forEach(function (ville) {
      makeCheckbox(villesEl, ville, ville, filterState.villes.has(ville), function (checked) {
        if (checked) filterState.villes.add(ville);
        else filterState.villes.delete(ville);
      });
    });

    const jours = Object.keys(data.jours || {}).sort(function (a, b) { return Number(a) - Number(b); });
    jours.forEach(function (jour) {
      const sheets = (data.jours[jour] || []).join(", ");
      makeCheckbox(joursEl, "Jour " + jour + (sheets ? " (" + sheets + ")" : ""), jour, filterState.jours.has(jour), function (checked) {
        if (checked) filterState.jours.add(jour);
        else filterState.jours.delete(jour);
      });
    });

    data.onglets.forEach(function (onglet) {
      makeCheckbox(ongletsEl, onglet, onglet, filterState.onglets.has(onglet), function (checked) {
        if (checked) filterState.onglets.add(onglet);
        else filterState.onglets.delete(onglet);
      });
    });
  }

  document.getElementById("btn-reset").addEventListener("click", function () {
    filterState.villes = new Set(data.villes);
    filterState.jours = new Set(Object.keys(data.jours || {}));
    filterState.onglets = new Set(data.onglets);
    applyPageFilter();

    document.querySelectorAll(".filter-group input[type=checkbox]").forEach(function (input) {
      input.checked = true;
    });
    if (pageFilter.mode === "ville" && pageFilter.ville) {
      document.querySelectorAll("#filter-villes input").forEach(function (input) {
        input.checked = input.id.includes(pageFilter.ville.toLowerCase().replace(/\s+/g, "-"));
      });
      filterState.villes = new Set([pageFilter.ville]);
    }
    refreshMarkers();
  });

  applyPageFilter();
  buildFilters();
  refreshMarkers();
})();
