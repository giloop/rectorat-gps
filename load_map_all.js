// Création de la map 
var map_0 = L.map("map_0", {
    center: [48.16056544872657, -1.0280],
    crs: L.CRS.EPSG3857,
    zoom: 8,
    zoomControl: true,
    preferCanvas: false,
    id: "my_map",
});

var tile_layer_1 = L.tileLayer(
    "https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
        "attribution": "Data by \u0026copy; \u003ca href=\"http://openstreetmap.org\"\u003eOpenStreetMap\u003c/a\u003e, under \u003ca href=\"http://www.openstreetmap.org/copyright\"\u003eODbL\u003c/a\u003e.",
        "detectRetina": false,
        "maxNativeZoom": 18,
        "maxZoom": 18,
        "minZoom": 1,
        "noWrap": false,
        "opacity": 1,
        "subdomains": "abc",
        "tms": false
    }
).addTo(map_0);

// Création des sous-groupes 
var groups = {
    '2022 KA121 SCH': null,
    '2021 KA120 SCH': null,
    '2022 KA122 SCH': null,
    '2022 KA131 HED': null,
    '2022 KA171 HED': null,
    'Autre': null
}

var arKeys = Object.keys(groups)

// Markers de type 
//map_0.addLayer(groups);

// Ajout des sous-groupes à la map
arKeys.forEach(el => {
    groups[el] = L.layerGroup().addTo(map_0);
})


function addMarker(item) {
    // Item : 
    // {"code": "code projet",
    // "organisme": row['Organisme candidat'],
    // "coordinates" : coords,
    // "couleur": couleur par code projet,
    // "infos": f"{info}""}

    if (groups.hasOwnProperty(item.code)) {
        myGroup = groups[item.code]
    } else {
        myGroup = groups['Autre']
    }

    var marker = L.marker(item.coordinates, {})
        .bindPopup(item.infos)
        .bindTooltip(item.infos, { "sticky": true })
        .setIcon(L.AwesomeMarkers.icon({
            "extraClasses": "fa-rotate-0",
            "icon": "ok-sign",
            "iconColor": "white",
            "markerColor": item.couleur,
            "prefix": "glyphicon"
        }));

    myGroup.addLayer(marker);
        
}

// Chargement des données
fetch('carto_programmes.json')
    .then(function(response) {
        return response.json()
    })
    .then(function(json) {
        // log des groupes
        const occurrences = json.reduce(function (acc, curr) {
            return acc[curr.code] ? ++acc[curr.code] : acc[curr.code] = 1, acc
          }, {});
          console.log(occurrences);
        // traitement du JSON
        json.forEach(el => { addMarker(el) });


    })
    .catch(function(error) {
        console.log('Il y a eu un problème avec l\'opération fetch: ' + error.message);
    });




var myLayers = {
    base_layers: {
        "openstreetmap": tile_layer_1,
    },
    overlays: groups,
};

L.control.layers(
    myLayers.base_layers,
    myLayers.overlays, {
        "autoZIndex": true,
        "collapsed": true,
        "position": "topright"
    }
).addTo(map_0);