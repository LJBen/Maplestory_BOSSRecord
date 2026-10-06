from typing import Any, Dict, List

import requests

from api_client import BASE_URL, HEADERS


def _as_int(value: Any) -> int:
	"""Convert an API numeric value to an integer, using 0 when unavailable."""
	try:
		return int(value)
	except (TypeError, ValueError):
		return 0


def _summarize_item_equipment(
	payload: Dict[str, Any],
) -> Dict[str, List[Dict[str, Any]]]:
	equipment = payload.get("item_equipment", [])
	if not isinstance(equipment, list):
		equipment = []

	restraint_ring_level = 0
	continuous_ring_level = 0
	weapon_name = ""
	soul_weapon_grade = 0
	has_heavenly_breath = False

	for item in equipment:
		if not isinstance(item, dict):
			continue

		item_name = str(item.get("item_name") or "")
		if "規範戒指" in item_name:
			restraint_ring_level = _as_int(item.get("special_ring_level"))
		elif "永續戒指" in item_name:
			continuous_ring_level = _as_int(item.get("special_ring_level"))

		if item.get("item_equipment_slot") == "武器":
			weapon_name = item_name
			soul_weapon_grade = _as_int(item.get("soul_weapon_grade"))

		if (
			item.get("item_equipment_part") == "戒指"
			and item_name == "天上的氣息"
		):
			has_heavenly_breath = True

	return {
		"item_equipment": [
			{
				"item_name": "規範戒指",
				"special_ring_level": restraint_ring_level,
			},
			{
				"item_name": "永續戒指",
				"special_ring_level": continuous_ring_level,
			},
			{
				"item_equipment_slot": "武器",
				"item_name": weapon_name,
				"soul_weapon_grade": soul_weapon_grade,
				"has_fate_or_genesis": "命運" in weapon_name or "創世" in weapon_name,
			},
			{
				"item_equipment_part": "戒指",
				"item_name": "天上的氣息",
				"equipped": has_heavenly_breath,
			},
		]
	}


def get_character_specific_info(ocid: str) -> Dict[str, List[Dict[str, Any]]]:
	"""Fetch and summarize the character's relevant item equipment."""
	url = f"{BASE_URL}/character/item-equipment"
	response = requests.get(url, headers=HEADERS, params={"ocid": ocid})
	response.raise_for_status()
	return _summarize_item_equipment(response.json())
