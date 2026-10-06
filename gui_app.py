import os
import customtkinter as ct
from tkinter import filedialog, messagebox
from api_client import get_character_full_data
from excel_writer import append_to_challenger_sheet
from get_character_specific_info import get_character_specific_info


class App(ct.CTk):
    def __init__(self):
        super().__init__()
        self.title("打王資料記錄")
        self.geometry("920x660")

        # Top: simple name input with submit button to the right
        top = ct.CTkFrame(self)
        top.pack(fill="x", padx=12, pady=12)

        ct.CTkLabel(top, text="角色名稱:").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.name_var = ct.StringVar()
        self.name_entry = ct.CTkEntry(top, textvariable=self.name_var, width=280, placeholder_text="最多10字")
        self.name_entry.grid(row=0, column=1, padx=6, pady=6)
        ct.CTkButton(top, text="送出", command=self.on_submit).grid(row=0, column=2, padx=6, pady=6)

        # Extra info panel: boss selection, clear time, categories, remarks
        self.extra_frame = ct.CTkFrame(self)

        ct.CTkLabel(self.extra_frame, text="難度:").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        difficulty_options = ["簡單", "普通", "困難", "極限"]
        self.difficulty_var = ct.StringVar(value=difficulty_options[0])
        ct.CTkOptionMenu(self.extra_frame, values=difficulty_options, variable=self.difficulty_var, width=120).grid(row=0, column=1, padx=6, pady=6, sticky="w")

        ct.CTkLabel(self.extra_frame, text="王:").grid(row=0, column=2, sticky="w", padx=6, pady=6)
        boss_options = [
            "史烏", "戴米安","綠水靈", "露希妲", "威爾",
            "戴斯克","真·希拉", "頓凱爾", "老黑"
            ,"賽連", "卡洛斯", "敵對者", "咖凌", "林波", "人馬","凱伊","凶星"
        ]
        self.boss_var = ct.StringVar(value=boss_options[0])
        ct.CTkOptionMenu(self.extra_frame, values=boss_options, variable=self.boss_var, width=120).grid(row=0, column=3, padx=6, pady=6, sticky="w")

        ct.CTkLabel(self.extra_frame, text="通關時間:").grid(row=0, column=4, sticky="w", padx=6, pady=6)
        self.clear_time_var = ct.StringVar()
        ct.CTkEntry(self.extra_frame, textvariable=self.clear_time_var, width=160, placeholder_text="HH:MM:SS").grid(row=0, column=5, padx=6, pady=6, sticky="w")

        self.checkbox_frame = ct.CTkFrame(self)

        self.category_vars = {
            "祕笈(紅)": ct.BooleanVar(),
            "祕笈(綠)": ct.BooleanVar(),
            "祕笈(橘)": ct.BooleanVar(),
            "創世": ct.BooleanVar(),
            "天上": ct.BooleanVar(),
            "挑戰者": ct.BooleanVar(),
            "一般服": ct.BooleanVar(),
        }
        col = 0
        for label, var in self.category_vars.items():
            ct.CTkCheckBox(self.checkbox_frame, text=label, variable=var).grid(row=0, column=col, padx=8, pady=6, sticky="w")
            col += 1

        self.remark_frame = ct.CTkFrame(self)
        ct.CTkLabel(self.remark_frame, text="規範:").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.spec_var = ct.StringVar()
        ct.CTkEntry(self.remark_frame, textvariable=self.spec_var, width=60, placeholder_text="數字").grid(row=0, column=1, padx=6, pady=6, sticky="w")
        ct.CTkLabel(self.remark_frame, text="永續:").grid(row=0, column=2, sticky="w", padx=6, pady=6)
        self.sustain_var = ct.StringVar()
        ct.CTkEntry(self.remark_frame, textvariable=self.sustain_var, width=60, placeholder_text="數字").grid(row=0, column=3, padx=6, pady=6, sticky="w")
        ct.CTkLabel(self.remark_frame, text="靈魂武器:").grid(row=0, column=4, sticky="w", padx=6, pady=6)
        self.soul_var = ct.StringVar()
        ct.CTkEntry(self.remark_frame, textvariable=self.soul_var, width=60, placeholder_text="輸入").grid(row=0, column=5, padx=6, pady=6, sticky="w")
        ct.CTkLabel(self.remark_frame, text="備註:").grid(row=0, column=6, sticky="w", padx=6, pady=6)
        self.note_var = ct.StringVar()
        ct.CTkEntry(self.remark_frame, textvariable=self.note_var, width=320, placeholder_text="一行備註").grid(row=0, column=7, padx=6, pady=6, sticky="w")

        self.upload_frame = ct.CTkFrame(self)
        self.upload_frame.grid_columnconfigure(0, weight=1)
        self.upload_button = ct.CTkButton(self.upload_frame, text="上傳", command=self.on_upload, width=120)
        self.upload_button.grid(row=0, column=0, sticky="e", padx=12, pady=12)

        # Reserved area below for API result fields
        self.fields_frame = ct.CTkFrame(self, height=320)
        self.fields_frame.pack(fill="both", expand=True, padx=12, pady=6)
        ct.CTkLabel(self.fields_frame, text="請輸入角色名稱後按送出以查詢資料").pack(padx=8, pady=8)

        self.extra_shown = False

        # keep placeholders for future functionality
        self.current_record = None
        self.entry_widgets = {}
        self.excel_path_var = ct.StringVar()


    def on_submit(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("缺少名稱", "請輸入角色名稱。")
            return
        if len(name) > 20:
            messagebox.showwarning("過長", "名稱超過20字，已截斷。")
            name = name[:20]
            self.name_var.set(name)
        # call API and display editable fields
        try:
            data = get_character_full_data(name)
            ocid_data = data.get("ocid_data") or {}
            ocid = ocid_data.get("ocid")
            if not ocid:
                raise ValueError("無法取得 OCID，請確認角色名稱是否正確。")
            specific_info = get_character_specific_info(ocid)
        except Exception as e:
            messagebox.showerror("API 錯誤", str(e))
            return

        item_equipment = specific_info.get("item_equipment", [])
        data["item_equipment"] = item_equipment
        restraint_ring = next(
            (item for item in item_equipment if item.get("item_name") == "規範戒指"),
            {},
        )
        continuous_ring = next(
            (item for item in item_equipment if item.get("item_name") == "永續戒指"),
            {},
        )
        weapon = next(
            (item for item in item_equipment if item.get("item_equipment_slot") == "武器"),
            {},
        )
        heavenly_breath = next(
            (
                item
                for item in item_equipment
                if item.get("item_equipment_part") == "戒指"
                and item.get("item_name") == "天上的氣息"
            ),
            {},
        )

        self.clear_time_var.set("")
        for label in ("祕笈(紅)", "祕笈(綠)", "祕笈(橘)"):
            self.category_vars[label].set(False)
        self.note_var.set("")
        self.spec_var.set(str(restraint_ring.get("special_ring_level", 0)))
        self.sustain_var.set(str(continuous_ring.get("special_ring_level", 0)))
        self.soul_var.set(str(weapon.get("soul_weapon_grade", 0)))
        self.category_vars["創世"].set(bool(weapon.get("has_fate_or_genesis", False)))
        self.category_vars["天上"].set(bool(heavenly_breath.get("equipped", False)))

        basic = data.get("basic", {})
        is_challenger_world = basic.get("world_name") == "挑戰者"
        self.category_vars["挑戰者"].set(is_challenger_world)
        self.category_vars["一般服"].set(not is_challenger_world)
        stat = data.get("stat", {})
        hexamatrix = data.get("hexamatrix", [])

        # prepare display values
        character_class = basic.get("character_class", "")
        character_level = basic.get("character_level", "")
        stat_value_raw = stat.get("stat_value", "")
        stat_display = self._format_stat(stat_value_raw)

        # clear and build fields
        for w in self.fields_frame.winfo_children():
            w.destroy()
        self.entry_widgets = {}

        # 職業
        ct.CTkLabel(self.fields_frame, text="職業:").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        class_var = ct.StringVar(value=character_class)
        ct.CTkEntry(self.fields_frame, textvariable=class_var, width=360).grid(row=0, column=1, padx=6, pady=6, sticky="w")
        self.entry_widgets["character_class"] = class_var

        # 等級
        ct.CTkLabel(self.fields_frame, text="等級:").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        level_var = ct.StringVar(value=str(character_level))
        ct.CTkEntry(self.fields_frame, textvariable=level_var, width=120).grid(row=1, column=1, padx=6, pady=6, sticky="w")
        self.entry_widgets["character_level"] = level_var

        # 戰鬥力
        ct.CTkLabel(self.fields_frame, text="戰鬥力:").grid(row=2, column=0, sticky="w", padx=6, pady=6)
        power_var = ct.StringVar(value=stat_display)
        ct.CTkEntry(self.fields_frame, textvariable=power_var, width=200).grid(row=2, column=1, padx=6, pady=6, sticky="w")
        # store raw as well
        self.entry_widgets["stat_value_display"] = power_var
        self.entry_widgets["stat_value_raw"] = stat_value_raw

        # Hexamatrix: split into 4 core types and show as editable entries
        # parse list like "技能核心: 1/0"
        core_map = { }
        for item in hexamatrix:
            if ":" in item:
                k, v = item.split(":", 1)
                core_map[k.strip()] = v.strip()

        # 技能核心
        ct.CTkLabel(self.fields_frame, text="技能核心:").grid(row=3, column=0, sticky="w", padx=6, pady=6)
        skill_var = ct.StringVar(value=core_map.get("技能核心", ""))
        ct.CTkEntry(self.fields_frame, textvariable=skill_var, width=360).grid(row=3, column=1, padx=6, pady=6, sticky="w")
        self.entry_widgets["hex_skill"] = skill_var

        # 精通核心
        ct.CTkLabel(self.fields_frame, text="精通核心:").grid(row=4, column=0, sticky="w", padx=6, pady=6)
        master_var = ct.StringVar(value=core_map.get("精通核心", ""))
        ct.CTkEntry(self.fields_frame, textvariable=master_var, width=360).grid(row=4, column=1, padx=6, pady=6, sticky="w")
        self.entry_widgets["hex_master"] = master_var

        # 強化核心
        ct.CTkLabel(self.fields_frame, text="強化核心:").grid(row=5, column=0, sticky="w", padx=6, pady=6)
        enhance_var = ct.StringVar(value=core_map.get("強化核心", ""))
        ct.CTkEntry(self.fields_frame, textvariable=enhance_var, width=360).grid(row=5, column=1, padx=6, pady=6, sticky="w")
        self.entry_widgets["hex_enhance"] = enhance_var

        # 共用核心
        ct.CTkLabel(self.fields_frame, text="共用核心:").grid(row=6, column=0, sticky="w", padx=6, pady=6)
        shared_var = ct.StringVar(value=core_map.get("共用核心", ""))
        ct.CTkEntry(self.fields_frame, textvariable=shared_var, width=360).grid(row=6, column=1, padx=6, pady=6, sticky="w")
        self.entry_widgets["hex_shared"] = shared_var

        # show the extra panels after query
        self.show_extra_panels()

        # keep current record for reference
        self.current_record = data

        # keep current record for reference
        self.current_record = data

    def _format_stat(self, value):
        # convert numeric string to '億'/'萬' units
        try:
            v = int(str(value))
        except Exception:
            return str(value)
        yi = v // 100000000
        rest = v % 100000000
        wan = rest // 10000
        if yi > 0:
            return f"{yi}億{wan}萬"
        if wan > 0:
            return f"{wan}萬"
        return str(v)

    def show_extra_panels(self):
        if not self.extra_shown:
            self.extra_frame.pack(fill="x", padx=12, pady=6)
            self.checkbox_frame.pack(fill="x", padx=12, pady=6)
            self.remark_frame.pack(fill="x", padx=12, pady=6)
            self.upload_frame.pack(fill="x", padx=12, pady=(0, 12))
            self.extra_shown = True

    def reset_form(self):
        self.clear_time_var.set("")

    def on_upload(self):
        # collect edited values
        character_class = self.entry_widgets.get("character_class").get()
        character_level = self.entry_widgets.get("character_level").get()
        stat_display = self.entry_widgets.get("stat_value_display").get()
        # collect split hexamatrix fields
        hex_skill = self.entry_widgets.get("hex_skill").get()
        hex_master = self.entry_widgets.get("hex_master").get()
        hex_enhance = self.entry_widgets.get("hex_enhance").get()
        hex_shared = self.entry_widgets.get("hex_shared").get()
        # combine into a readable representation
        hex_text = f"技能核心: {hex_skill}\n精通核心: {hex_master}\n強化核心: {hex_enhance}\n共用核心: {hex_shared}"

        difficulty = self.difficulty_var.get()
        boss = self.boss_var.get()
        clear_time = self.clear_time_var.get().strip()
        categories = {label: bool(var.get()) for label, var in self.category_vars.items()}
        spec_value = self.spec_var.get().strip()
        sustain_value = self.sustain_var.get().strip()
        soul_value = self.soul_var.get().strip()
        note = self.note_var.get().strip()

        gui_data = {
            "character_class": character_class,
            "character_level": character_level,
            "stat_value_display": stat_display,
            "difficulty": difficulty,
            "boss": boss,
            "clear_time": clear_time,
            "hex_skill": hex_skill,
            "hex_master": hex_master,
            "hex_enhance": hex_enhance,
            "hex_shared": hex_shared,
            "祕笈(紅)": categories.get("祕笈(紅)", False),
            "祕笈(綠)": categories.get("祕笈(綠)", False),
            "祕笈(橘)": categories.get("祕笈(橘)", False),
            "創世": categories.get("創世", False),
            "天上": categories.get("天上", False),
            "挑戰者": categories.get("挑戰者", False),
            "一般服": categories.get("一般服", False),
            "規範": spec_value,
            "永續": sustain_value,
            "靈魂武器": soul_value,
            "note": note,
        }

        excel_path = os.path.join(os.path.dirname(__file__), "BOSS.xlsm")
        try:
            append_to_challenger_sheet(excel_path, gui_data)
        except Exception as exc:
            messagebox.showerror("寫入 Excel 失敗", str(exc))
            return

        self.current_record = gui_data
        self.reset_form()
        messagebox.showinfo("已上傳", "資料已寫入 Excel。")

def run():
    ct.set_appearance_mode("System")
    ct.set_default_color_theme("blue")
    app = App()
    app.mainloop()
