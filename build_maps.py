"""Download cached OpenStreetMap tiles and render itinerary map graphics."""

from __future__ import annotations

import argparse
import math
import urllib.request
from pathlib import Path

from PIL import Image, ImageDraw, ImageEnhance, ImageFont

from map_data import MAPS


BASE = Path(__file__).resolve().parent
CACHE = BASE / "map_cache" / "osm"
TILE_SIZE = 256
USER_AGENT = "vacation-research-2026/1.0 (personal itinerary map builder)"
TILE_SOURCES = {
    "osm": {
        "url": "https://tile.openstreetmap.org/{z}/{x}/{y}.png",
        "attribution": "© OpenStreetMap contributors",
    },
    "carto_positron": {
        "url": "https://a.basemaps.cartocdn.com/rastertiles/light_all/{z}/{x}/{y}.png",
        "attribution": "© OpenStreetMap contributors · © CARTO",
    },
    "carto_voyager": {
        "url": "https://a.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}.png",
        "attribution": "© OpenStreetMap contributors · © CARTO",
    },
    "esri_world_street": {
        "url": "https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}",
        "attribution": "Sources: Esri, HERE, Garmin, USGS, OSM contributors",
    },
}


def world_pixel(lat: float, lon: float, zoom: int) -> tuple[float, float]:
    scale = TILE_SIZE * 2**zoom
    x = (lon + 180.0) / 360.0 * scale
    lat_rad = math.radians(lat)
    y = (1.0 - math.asinh(math.tan(lat_rad)) / math.pi) / 2.0 * scale
    return x, y


def tile_path(source: str, zoom: int, x: int, y: int) -> Path:
    source_cache = CACHE if source == "osm" else CACHE.parent / source
    return source_cache / str(zoom) / str(x) / f"{y}.png"


def get_tile(source: str, zoom: int, x: int, y: int, *, offline: bool) -> Image.Image:
    path = tile_path(source, zoom, x, y)
    if not path.exists():
        if offline:
            raise FileNotFoundError(f"Missing cached tile in offline mode: {path}")
        path.parent.mkdir(parents=True, exist_ok=True)
        url = TILE_SOURCES[source]["url"].format(z=zoom, x=x, y=y)
        request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
        with urllib.request.urlopen(request, timeout=20) as response:
            path.write_bytes(response.read())
    with Image.open(path) as tile:
        return tile.convert("RGB")


def font(size: int, *, bold: bool = False) -> ImageFont.ImageFont:
    filename = "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf"
    path = Path("/usr/share/fonts/truetype/dejavu") / filename
    return ImageFont.truetype(str(path), size) if path.exists() else ImageFont.load_default()


def draw_pin(draw: ImageDraw.ImageDraw, x: float, y: float, number: int, scale: float) -> None:
    radius = int(12 * scale)
    stem = int(10 * scale)
    center_y = int(y - radius - stem)
    outline = max(2, int(2 * scale))
    draw.line((x, center_y + radius - 1, x, y), fill="white", width=outline + 3)
    draw.line((x, center_y + radius - 1, x, y), fill="#c53b32", width=outline)
    draw.ellipse(
        (x - radius, center_y - radius, x + radius, center_y + radius),
        fill="#c53b32",
        outline="white",
        width=outline,
    )
    label_font = font(int(11 * scale), bold=True)
    text = str(number)
    box = draw.textbbox((0, 0), text, font=label_font)
    draw.text(
        (x - (box[2] - box[0]) / 2, center_y - (box[3] - box[1]) / 2 - box[1]),
        text,
        fill="white",
        font=label_font,
    )


def render_map(map_id: str, *, offline: bool = False) -> Path:
    spec = MAPS[map_id]
    tile_source = spec.get("tile_source", "osm")
    zoom = spec["zoom"]
    viewport_width, viewport_height = spec["viewport"]
    output_width, output_height = spec["output_size"]
    points = [world_pixel(place["lat"], place["lon"], zoom) for place in spec["places"]]
    center_x = (min(x for x, _ in points) + max(x for x, _ in points)) / 2
    center_y = (min(y for _, y in points) + max(y for _, y in points)) / 2
    offset_x, offset_y = spec.get("center_offset", (0, 0))
    center_x += offset_x
    center_y += offset_y
    left = center_x - viewport_width / 2
    top = center_y - viewport_height / 2
    right = left + viewport_width
    bottom = top + viewport_height

    min_tile_x, max_tile_x = math.floor(left / TILE_SIZE), math.floor((right - 1) / TILE_SIZE)
    min_tile_y, max_tile_y = math.floor(top / TILE_SIZE), math.floor((bottom - 1) / TILE_SIZE)
    mosaic = Image.new(
        "RGB",
        ((max_tile_x - min_tile_x + 1) * TILE_SIZE, (max_tile_y - min_tile_y + 1) * TILE_SIZE),
    )
    for tile_y in range(min_tile_y, max_tile_y + 1):
        for tile_x in range(min_tile_x, max_tile_x + 1):
            tile = get_tile(tile_source, zoom, tile_x, tile_y, offline=offline)
            mosaic.paste(tile, ((tile_x - min_tile_x) * TILE_SIZE, (tile_y - min_tile_y) * TILE_SIZE))

    crop_left = int(round(left - min_tile_x * TILE_SIZE))
    crop_top = int(round(top - min_tile_y * TILE_SIZE))
    image = mosaic.crop((crop_left, crop_top, crop_left + viewport_width, crop_top + viewport_height))
    image = image.resize((output_width, output_height), Image.Resampling.LANCZOS)
    tile_adjust = spec.get("tile_adjust", {})
    if "brightness" in tile_adjust:
        image = ImageEnhance.Brightness(image).enhance(tile_adjust["brightness"])
    if "contrast" in tile_adjust:
        image = ImageEnhance.Contrast(image).enhance(tile_adjust["contrast"])
    if "color" in tile_adjust:
        image = ImageEnhance.Color(image).enhance(tile_adjust["color"])
    scale_x, scale_y = output_width / viewport_width, output_height / viewport_height
    draw = ImageDraw.Draw(image)
    route = spec.get("route")
    if route:
        route_pixels = [
            ((x - left) * scale_x, (y - top) * scale_y)
            for x, y in (world_pixel(lat, lon, zoom) for lat, lon in route)
        ]
        draw.line(route_pixels, fill="white", width=max(7, int(5 * scale_x)), joint="curve")
        draw.line(route_pixels, fill="#2f6f9f", width=max(3, int(2.5 * scale_x)), joint="curve")
    for place, (point_x, point_y) in zip(spec["places"], points):
        draw_pin(draw, (point_x - left) * scale_x, (point_y - top) * scale_y, place["number"], scale_x)

    attribution = TILE_SOURCES[tile_source]["attribution"]
    attribution_font = font(13)
    box = draw.textbbox((0, 0), attribution, font=attribution_font)
    padding = 4
    x = output_width - (box[2] - box[0]) - padding * 2
    y = output_height - (box[3] - box[1]) - padding * 2
    draw.rectangle((x, y, output_width, output_height), fill=(255, 255, 255, 220))
    draw.text((x + padding, y + padding - box[1]), attribution, fill="#334", font=attribution_font)

    output = BASE / spec["output"]
    output.parent.mkdir(parents=True, exist_ok=True)
    image.save(output, optimize=True)
    return output


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("map_ids", nargs="*", choices=sorted(MAPS), help="Maps to build; defaults to all")
    parser.add_argument("--offline", action="store_true", help="Fail rather than downloading missing tiles")
    args = parser.parse_args()
    map_ids = args.map_ids or list(MAPS)
    for map_id in map_ids:
        output = render_map(map_id, offline=args.offline)
        print(f"Wrote {output.relative_to(BASE)}")


if __name__ == "__main__":
    main()
