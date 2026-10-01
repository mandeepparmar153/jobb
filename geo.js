/* Mod Entreprenad GPS – delade geofunktioner (används av arbetarsidan och adminsidan). */
(function (root) {
  'use strict';
  var R = 6371000;
  var rad = Math.PI / 180;

  function dist(lat1, lng1, lat2, lng2) {
    var dLat = (lat2 - lat1) * rad, dLng = (lng2 - lng1) * rad;
    var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
      Math.cos(lat1 * rad) * Math.cos(lat2 * rad) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
    return 2 * R * Math.asin(Math.min(1, Math.sqrt(a)));
  }

  // Lokal platt projektion i meter runt (lat0, lng0)
  function projector(lat0, lng0) {
    var kx = Math.cos(lat0 * rad) * rad * R, ky = rad * R;
    return {
      to: function (lat, lng) { return [(lng - lng0) * kx, (lat - lat0) * ky]; },
      from: function (x, y) { return [lat0 + y / ky, lng0 + x / kx]; },
    };
  }

  function pointInPolygon(lat, lng, poly) {
    var inside = false;
    for (var i = 0, j = poly.length - 1; i < poly.length; j = i++) {
      var yi = poly[i][0], xi = poly[i][1], yj = poly[j][0], xj = poly[j][1];
      if (((yi > lat) !== (yj > lat)) && (lng < (xj - xi) * (lat - yi) / (yj - yi) + xi)) inside = !inside;
    }
    return inside;
  }

  function segDist(px, py, ax, ay, bx, by) {
    var dx = bx - ax, dy = by - ay, l2 = dx * dx + dy * dy;
    var t = l2 ? Math.max(0, Math.min(1, ((px - ax) * dx + (py - ay) * dy) / l2)) : 0;
    var x = ax + t * dx - px, y = ay + t * dy - py;
    return Math.sqrt(x * x + y * y);
  }

  function hasPoly(site) { return site && site.polygon && site.polygon.length >= 3; }

  // Avstånd i meter från punkt till plats (0 om innanför)
  function distToSite(lat, lng, site) {
    if (hasPoly(site)) {
      if (pointInPolygon(lat, lng, site.polygon)) return 0;
      var pr = projector(lat, lng), best = Infinity, p = site.polygon;
      for (var i = 0, j = p.length - 1; i < p.length; j = i++) {
        var a = pr.to(p[j][0], p[j][1]), b = pr.to(p[i][0], p[i][1]);
        best = Math.min(best, segDist(0, 0, a[0], a[1], b[0], b[1]));
      }
      return best;
    }
    return Math.max(0, dist(lat, lng, site.lat, site.lng) - (site.radius || 50));
  }

  function inSite(lat, lng, site, margin) { return distToSite(lat, lng, site) <= (margin || 0); }

  function nearestSite(lat, lng, sites, maxMeters) {
    var best = null, bd = Infinity;
    (sites || []).forEach(function (s) {
      var d = distToSite(lat, lng, s);
      if (d < bd) { bd = d; best = s; }
    });
    return best && bd <= (maxMeters == null ? Infinity : maxMeters) ? { site: best, meters: bd } : null;
  }

  /**
   * Hitta stopp. points: [[ts, lat, lng, acc], ...] sorterade på tid.
   * Ett stopp = minst minMinutes inom radius meter från klustrets mittpunkt.
   */
  function detectStops(points, opt) {
    opt = opt || {};
    var radius = opt.radius || 50, minMs = (opt.minMinutes || 5) * 60000, gapMs = (opt.gapMinutes || 10) * 60000;
    var stops = [], n = points.length, i = 0;
    while (i < n) {
      var sLat = points[i][1], sLng = points[i][2], cnt = 1, j = i, maxGap = 0;
      for (var k = i + 1; k < n; k++) {
        var cLat = sLat / cnt, cLng = sLng / cnt;
        if (dist(points[k][1], points[k][2], cLat, cLng) > radius) break;
        maxGap = Math.max(maxGap, points[k][0] - points[k - 1][0]);
        sLat += points[k][1]; sLng += points[k][2]; cnt++; j = k;
      }
      var dur = points[j][0] - points[i][0];
      if (dur >= minMs) {
        stops.push({ start: points[i][0], end: points[j][0], lat: sLat / cnt, lng: sLng / cnt, minutes: Math.round(dur / 60000), n: cnt, gap: maxGap > gapMs });
        i = j + 1;
      } else {
        i++;
      }
    }
    return stops;
  }

  // Luckor: tid mellan två punkter längre än gapMinutes
  function detectGaps(points, gapMinutes) {
    var g = (gapMinutes || 10) * 60000, out = [];
    for (var i = 1; i < points.length; i++) {
      if (points[i][0] - points[i - 1][0] > g) out.push({ from: points[i - 1][0], to: points[i][0], minutes: Math.round((points[i][0] - points[i - 1][0]) / 60000) });
    }
    return out;
  }

  /**
   * Ungefär hur stor del (%) av ett område som spåret har täckt.
   * widthM = halva arbetsbredden (t.ex. 4 m för plog/traktor).
   */
  function coverage(polygon, tracks, widthM) {
    if (!polygon || polygon.length < 3 || !tracks || !tracks.length) return 0;
    if (typeof tracks[0][0] === 'number') tracks = [tracks];
    widthM = widthM || 4;
    var lat0 = 0, lng0 = 0;
    polygon.forEach(function (p) { lat0 += p[0]; lng0 += p[1]; });
    lat0 /= polygon.length; lng0 /= polygon.length;
    var pr = projector(lat0, lng0);
    var poly = polygon.map(function (p) { return pr.to(p[0], p[1]); });
    var minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    poly.forEach(function (p) { minX = Math.min(minX, p[0]); maxX = Math.max(maxX, p[0]); minY = Math.min(minY, p[1]); maxY = Math.max(maxY, p[1]); });
    var area = (maxX - minX) * (maxY - minY);
    var cell = Math.max(2, Math.sqrt(area / 4000));
    var segs = [];
    tracks.forEach(function (points) {
      for (var i = 1; i < points.length; i++) {
        var a = points[i - 1], b = points[i];
        if (b[0] - a[0] > 3 * 60000 || dist(a[1], a[2], b[1], b[2]) > 300) continue;
        var pa = pr.to(a[1], a[2]), pb = pr.to(b[1], b[2]);
        if (Math.max(pa[0], pb[0]) < minX - widthM || Math.min(pa[0], pb[0]) > maxX + widthM ||
            Math.max(pa[1], pb[1]) < minY - widthM || Math.min(pa[1], pb[1]) > maxY + widthM) continue;
        segs.push([pa[0], pa[1], pb[0], pb[1]]);
      }
      // Enskilda punkter räknas också (stillastående maskin)
      points.forEach(function (p) {
        var q = pr.to(p[1], p[2]);
        if (q[0] < minX - widthM || q[0] > maxX + widthM || q[1] < minY - widthM || q[1] > maxY + widthM) return;
        segs.push([q[0], q[1], q[0], q[1]]);
      });
    });
    var inside = 0, covered = 0;
    var polyLL = polygon;
    for (var x = minX + cell / 2; x < maxX; x += cell) {
      for (var y = minY + cell / 2; y < maxY; y += cell) {
        var ll = pr.from(x, y);
        if (!pointInPolygon(ll[0], ll[1], polyLL)) continue;
        inside++;
        for (var s = 0; s < segs.length; s++) {
          var sg = segs[s];
          if (segDist(x, y, sg[0], sg[1], sg[2], sg[3]) <= widthM + cell / 2) { covered++; break; }
        }
      }
    }
    return inside ? Math.round(100 * covered / inside) : 0;
  }

  // Minuter som en användares stopp har legat på en plats
  function minutesAtSite(stops, site, margin) {
    var m = 0;
    stops.forEach(function (st) { if (inSite(st.lat, st.lng, site, margin == null ? 25 : margin)) m += st.minutes; });
    return m;
  }

  var api = {
    dist: dist, pointInPolygon: pointInPolygon, distToSite: distToSite, inSite: inSite,
    nearestSite: nearestSite, detectStops: detectStops, detectGaps: detectGaps,
    coverage: coverage, minutesAtSite: minutesAtSite, projector: projector,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Geo = api;
})(this);
