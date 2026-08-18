from collections import defaultdict

import requests

API_KEY = "test_16c00444df3cd677fcb60dde233803328d270c30e10d8bb8b4cd19e143fb81eaefe8d04e6d233bd35cf2fabdeb93fb0d"
BASE_URL = "https://open.api.nexon.com/maplestorytw/v1"
HEADERS = {
    "x-nxopen-api-key": API_KEY,
}


def get_character_ocid(character_name: str) -> dict:
    """Get the character OCID by character name."""
    url = f"{BASE_URL}/id"
    params = {"character_name": character_name}
    response = requests.get(url, headers=HEADERS, params=params)
    response.raise_for_status()
    return response.json()


def get_character_basic(ocid: str) -> dict:
    """Get basic character info including class and level."""
    url = f"{BASE_URL}/character/basic"
    params = {"ocid": ocid}
    response = requests.get(url, headers=HEADERS, params=params)
    response.raise_for_status()
    data = response.json()
    return {
        "character_class": data.get("character_class"),
        "character_level": data.get("character_level"),
    }


def get_character_stat(ocid: str) -> dict:
    """Get character stats such as final stat / combat power."""
    url = f"{BASE_URL}/character/stat"
    params = {"ocid": ocid}
    response = requests.get(url, headers=HEADERS, params=params)
    response.raise_for_status()
    data = response.json()
    return {
        "stat_value": data.get("final_stat")[-2].get("stat_value"),
    }

def get_character_hexamatrix(ocid: str) -> dict:
    """Get character Hexa Matrix info."""
    url = f"{BASE_URL}/character/hexamatrix"
    params = {"ocid": ocid}
    response = requests.get(url, headers=HEADERS, params=params)
    response.raise_for_status()
    # 建立分類用的字典
    type_groups = defaultdict(list)
    data = response.json()

    # 遍歷資料進行分類
    for core in data["character_hexa_core_equipment"]:
        
        core_type = core["hexa_core_type"]
        core_level = str(core["hexa_core_level"])

        type_groups[core_type].append(core_level)

    result = []

    for core_type, values in type_groups.items():

        # 用 '/' 將陣列裡面的東西串起來
        skill = "/".join(values)
        result.append(f"{core_type}: {skill}")

    return result


def get_character_full_data(character_name: str) -> dict:
    """Get all Nexon data for a character by name."""
    ocid_data = get_character_ocid(character_name)
    ocid = ocid_data.get("ocid")
    if not ocid:
        raise ValueError("無法取得 OCID，請確認角色名稱是否正確。")

    basic = get_character_basic(ocid)
    stat = get_character_stat(ocid)
    hexamatrix = get_character_hexamatrix(ocid)

    return {
        "ocid_data": ocid_data,
        "basic": basic,
        "stat": stat,
        "hexamatrix": hexamatrix,
    }
