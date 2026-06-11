(function () {
  "use strict";

  var data = window.INSPECT_DATA;
  if (!data) {
    return;
  }

  var root = document.getElementById("inspect-root");
  var generated = document.getElementById("inspect-generated");
  var selectedDay = null;
  var map = null;

  var STATUS_LABELS = {
    ok: "OK",
    warn: "Attention",
    error: "Erreur",
  };

  function esc(text) {
    var div = document.createElement("div");
    div.textContent = text == null ? "" : String(text);
    return div.innerHTML;
  }

  function formatEuro(value) {
    if (value == null || value === 0) {
      return "—";
    }
    return Number(value).toFixed(2).replace(".", ",") + " €";
  }

  function formatDate(iso) {
    if (!iso) {
      return "—";
    }
    var parts = iso.split("-");
    if (parts.length !== 3) {
      return iso;
    }
    return parts[2] + "/" + parts[1] + "/" + parts[0];
  }

  function statusClass(status) {
    return status === "ok" ? "inspect-coverage-ok" : status === "warn" ? "inspect-coverage-warn" : "inspect-coverage-error";
  }

  function renderOverall() {
    var label = data.overall_status === "ok" ? "Pipeline cohérent" : data.overall_status === "warn" ? "Vérifications avec réserves" : "Incohérences détectées";
    return (
      '<div class="inspect-overall ' + esc(data.overall_status) + '">' +
      "<span>" + esc(label) + "</span>" +
      "<span>· " + esc(data.summary.activities) + " activités · " + esc(data.summary.on_map) + " sur la carte · " + esc(formatEuro(data.summary.budget)) + "</span>" +
      "</div>"
    );
  }

  function renderChecks() {
    var html = '<section class="inspect-section"><h2>État de santé</h2><div class="inspect-checks">';
    data.checks.forEach(function (check) {
      html +=
        '<article class="inspect-check">' +
        '<span class="inspect-check-badge ' + esc(check.status) + '">' + esc(STATUS_LABELS[check.status] || check.status) + "</span>" +
        '<p class="inspect-check-message">' + esc(check.message) + "</p>" +
        (check.details ? '<p class="inspect-check-details">' + esc(check.details) + "</p>" : "") +
        "</article>";
    });
    html += "</div></section>";
    return html;
  }

  function dayActivities(jour) {
    return data.activities.filter(function (item) {
      return item.jour === jour;
    });
  }

  function renderSidePanel() {
    if (!selectedDay) {
      return '<div class="inspect-side-panel"><h3>Détail du jour</h3><p>Cliquez sur un jour pour voir ses activités.</p></div>';
    }

    var day = data.days.find(function (item) {
      return item.jour === selectedDay;
    });
    var activities = dayActivities(selectedDay);
    var html =
      '<div class="inspect-side-panel">' +
      "<h3>Jour " + esc(selectedDay) + (day && day.date ? " · " + esc(formatDate(day.date)) : "") + "</h3>";

    if (day) {
      html += "<p><strong>" + esc(day.ville || "Ville non renseignée") + "</strong></p>";
      if (day.resume) {
        html += "<p>" + esc(day.resume) + "</p>";
      }
      html +=
        "<p>" + esc(day.activities) + " activité(s) · " + esc(day.geocoded) + " géoloc. · " + esc(formatEuro(day.prix)) + "</p>";
    }

    html += '<ul class="inspect-side-list">';
    activities.forEach(function (item) {
      var flags = [];
      if (item.on_map) {
        flags.push("carte");
      }
      if (item.in_overview) {
        flags.push("overview");
      }
      if (item.is_trajet) {
        flags.push("trajet");
      }
      html +=
        "<li><strong>" + esc(item.ordre) + "</strong> " + esc(item.nom) +
        (flags.length ? " <em>(" + esc(flags.join(", ")) + ")</em>" : "") +
        "</li>";
    });
    html += "</ul></div>";
    return html;
  }

  function renderTimeline() {
    var html =
      '<section class="inspect-section inspect-layout">' +
      '<div><h2>Frise chronologique</h2><div class="inspect-timeline" id="inspect-timeline">';

    data.days.forEach(function (day) {
      html +=
        '<button type="button" class="inspect-day' + (selectedDay === day.jour ? " is-selected" : "") + '" data-jour="' + esc(day.jour) + '">' +
        '<span class="inspect-day-badge" style="background:' + esc(day.couleur) + '">' + esc(day.jour) + "</span>" +
        '<div class="inspect-day-body"><h3>' + esc(day.ville || "—") + " · " + esc(formatDate(day.date)) + "</h3>" +
        '<p class="inspect-day-meta">' + esc(day.resume || "Pas de résumé overview") + "</p></div>" +
        '<div class="inspect-day-stats">' + esc(day.activities) + " act. · " + esc(day.foot_km) + " km pied</div>" +
        "</button>";
    });

    html += '</div></div><div id="inspect-side-panel">' + renderSidePanel() + "</div></section>";
    return html;
  }

  function renderCoverage() {
    var rows = data.coverage.filter(function (row) {
      return row.status !== "ok";
    });
    if (!rows.length) {
      rows = data.coverage.slice(0, 12);
    }

    var html =
      '<section class="inspect-section"><h2>Couverture des activités</h2>' +
      '<p class="stats-note">Activités hors carte, trajets ou jours absents de l\'overview sont signalés.</p>' +
      '<div class="stats-table-wrap"><table class="stats-table"><thead><tr>' +
      "<th>Étape</th><th>Lieu</th><th>Stats</th><th>Carte</th><th>Overview</th><th>État</th>" +
      "</tr></thead><tbody>";

    rows.forEach(function (row) {
      var overviewCell = data.has_overview ? (row.in_overview ? "✓" : "✗") : "—";
      html +=
        "<tr>" +
        "<td>J" + esc(row.jour) + " " + esc(row.ordre) + "</td>" +
        "<td>" + esc(row.nom) + "</td>" +
        "<td>" + (row.in_stats ? "✓" : "✗") + "</td>" +
        "<td>" + (row.on_map ? "✓" : "—") + "</td>" +
        "<td>" + overviewCell + "</td>" +
        '<td class="' + statusClass(row.status) + '">' + esc(STATUS_LABELS[row.status] || row.status) + "</td>" +
        "</tr>";
    });

    html += "</tbody></table></div></section>";
    return html;
  }

  function renderSources() {
    var html = '<section class="inspect-section"><h2>Fichiers sources</h2><div class="inspect-sources">';
    Object.keys(data.sources).forEach(function (key) {
      var source = data.sources[key];
      html +=
        '<div class="inspect-source' + (source.present ? "" : " missing") + '">' +
        "<strong>" + esc(key) + "</strong>" +
        (source.present ? esc(source.path) + "<br>" + esc(source.mtime || "") : "Fichier absent") +
        "</div>";
    });
    html += "</div></section>";
    return html;
  }

  function renderMap() {
    return '<section class="inspect-section"><h2>Carte de contrôle</h2><div id="inspect-map" class="inspect-map"></div></section>';
  }

  function renderJsonExplorer() {
    return (
      '<section class="inspect-section inspect-json">' +
      "<details><summary>Données brutes (inspect.json)</summary>" +
      "<pre>" + esc(JSON.stringify(data, null, 2)) + "</pre>" +
      "</details></section>"
    );
  }

  function bindTimeline() {
    var timeline = document.getElementById("inspect-timeline");
    if (!timeline) {
      return;
    }
    timeline.addEventListener("click", function (event) {
      var button = event.target.closest(".inspect-day");
      if (!button) {
        return;
      }
      selectedDay = Number(button.getAttribute("data-jour"));
      document.getElementById("inspect-side-panel").innerHTML = renderSidePanel();
      timeline.querySelectorAll(".inspect-day").forEach(function (node) {
        node.classList.toggle("is-selected", Number(node.getAttribute("data-jour")) === selectedDay);
      });
      highlightMapDay(selectedDay);
    });
  }

  function initMap() {
    var container = document.getElementById("inspect-map");
    if (!container || !window.L || !data.map_points.length) {
      return;
    }

    map = L.map(container, { scrollWheelZoom: false });
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      attribution: "&copy; OpenStreetMap",
      maxZoom: 18,
    }).addTo(map);

    var bounds = [];
    data.map_points.forEach(function (point) {
      var marker = L.circleMarker([point.lat, point.lon], {
        radius: 7,
        color: "#fff",
        weight: 2,
        fillColor: point.couleur,
        fillOpacity: 0.95,
      });
      marker.bindPopup("<strong>" + esc(point.ordre) + "</strong> " + esc(point.nom));
      marker.addTo(map);
      marker._inspectJour = point.jour;
      bounds.push([point.lat, point.lon]);
    });

    if (bounds.length) {
      map.fitBounds(bounds, { padding: [24, 24] });
    }
  }

  function highlightMapDay(jour) {
    if (!map) {
      return;
    }
    map.eachLayer(function (layer) {
      if (!layer._inspectJour) {
        return;
      }
      var active = layer._inspectJour === jour;
      layer.setStyle({
        radius: active ? 10 : 7,
        fillOpacity: active ? 1 : 0.55,
      });
      if (active) {
        layer.bringToFront();
      }
    });
  }

  function render() {
    if (generated) {
      generated.textContent = "Généré le " + data.generated_at + " · " + data.summary.title;
    }
    root.innerHTML =
      renderOverall() +
      renderChecks() +
      renderTimeline() +
      renderMap() +
      renderCoverage() +
      renderSources() +
      renderJsonExplorer();
    bindTimeline();
    initMap();
  }

  render();
})();
