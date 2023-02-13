import logger

log = logger.setup_applevel_logger(file_name='debug.log')
log.info('Log file created')
log.info("started main")

import json
import asyncio
import sys
import pyautogui
import rotation
import customtkinter as ctk
from tkinter import filedialog, TclError
from models import get_profiles, get_repeat_delay, config_blueprint, drm, get_config, get_spell, get_spec, add_spell
from models import add_spec
import ahk_converter
import webbrowser
import win32api
from pywinauto import keyboard
from time import sleep
import win32gui
from virtual_key_codes import VK_CODE as vkc

ctk.set_appearance_mode("System")  # Modes: system (default), light, dark
ctk.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green

app = ctk.CTk()  # create CTk window like you do with the Tk window
app.geometry("250x410")
app.resizable(False, False)
app.title("Parse God")

# noinspection PyTypeChecker
app.iconbitmap("icon.ico")
app.configure()


def get_resolution():
    desktop = win32gui.GetDesktopWindow()
    resolution = win32gui.GetWindowRect(desktop)
    profile = str(resolution[3]) + "p"
    return profile


def checkColor(x, y):
    rgb = [pyautogui.pixel(x, y)]
    return rgb


def update_toggle():
    try:
        if toggle.get() == 1:
            if win32api.GetKeyState(rotation.rotation_hotkey() == 1):
                keyboard.send_keys(rotation.rotation_hotkey())
            rotation.toggle = True
            switch_1.select()
            switch_1.configure(text="AFK Mode: ON")
        else:
            rotation.toggle = False
            switch_1.deselect()
            switch_1.configure(text="AFK Mode: OFF")
    except NameError:
        return "AFK Mode: OFF"


def save_and_close(window, value: str = None) -> None:
    print(value)
    window.destroy()


def confirmation_box(title: str, message: str, callback) -> None:
    confirmation = ctk.CTkToplevel(app)
    confirmation.title(title)
    ctk.CTkLabel(confirmation, text=message).grid(row=0, column=0, columnspan=2)
    yes_button = ctk.CTkButton(confirmation, text="Yes!", command=lambda: [callback(), confirmation.destroy()])
    yes_button.grid(row=1, column=0, pady=10, padx=10)
    no_button = ctk.CTkButton(confirmation, text="No!", command=lambda: confirmation.destroy())
    no_button.grid(row=1, column=1, pady=10, padx=10)
    app.wait_window(confirmation)


def input_box(title: str, text: str) -> str:
    dialog = ctk.CTkInputDialog(app, title=title, text=text)
    answer = dialog.get_input()
    return answer


def import_ahk() -> None:
    file = filedialog.askopenfilename(initialdir="/", title="Select file", filetypes=(("ahk files", "*.ahk"),
                                                                                      ("all files", "*.*")))
    print(file)
    ahk_converter.main(file, config_blueprint(current_profile=get_resolution()))


def set_profile(value: str) -> None:
    for profile in enumerate(get_profiles()):
        try:
            if profile[1] == value:
                config = get_config()
                config["current_profile"] = value
                save_config(config)
        except AttributeError:
            pass


def del_profile() -> None:
    config = get_config()
    for profile in enumerate(get_profiles()):
        if profile[1] == profiles.get():
            config["profiles"].pop(profile[0])
            if profile[1] == config["current_profile"] and len(config["profiles"]) > 1:
                if profile[0] == 0:
                    config["current_profile"] = get_profiles()[1]
            else:
                config["current_profile"] = ""
            save_config(config)
    sleep(0.1)
    profiles.configure(values=get_profiles())
    try:
        profiles.set(profiles.values[0])
        set_profile(profiles.values[0])
    except IndexError:
        profiles.set("")
        button_del.configure(state="disabled")


# TODO: Add inputboxes for the remaining settings of the config
def add_profile() -> str:
    profile_name = input_box("Add Profile", "Enter Profile name:")
    while profile_name is None:
        profile_name = input_box("Add Profile", "Enter Profile name:")
        continue
    config = get_config()
    try:
        config["profiles"].append(config_blueprint(profile_name)["profiles"][0])
        save_config(config)
    except AttributeError:
        pass
    except TypeError:
        config = config_blueprint(profile_name)
        save_config(config)
    try:
        profiles.configure(values=get_profiles())
        profiles.set(profile_name)
        button_del.configure(state="normal")
        set_profile(profile_name)
    except NameError:
        pass
    except AttributeError:
        pass
    return profile_name


def repeat_delay(time) -> None:
    config = get_config()
    delay_label.configure(text=f"Delay between spells (ms): {int(time)}")
    for profile in enumerate(get_profiles()):
        if profile[1] == profiles.get():
            config["profiles"][profile[0]]["repeat_delay"] = time / 1000
            save_config(config)


def update_profiles() -> list:
    profiles_entry = get_profiles()
    if profiles_entry:
        return profiles_entry
    else:
        # return [add_profile()]
        confirmation_box("No Profiles", "You have no profiles, would you like to import from AHK?", import_ahk)
        if get_profiles():
            return get_profiles()
        else:
            return [add_profile()]


def save_config(config):
    with open("config.json", "w") as write:
        json.dump(config, write, indent=2)


# Extra frame for more settings

class OptionsMenu:
    def __init__(self, header_name="OptionsMenu"):
        super().__init__()

        self.config = get_config()
        self.header_name = header_name
        self.drm_key = drm()
        self.options = ctk.CTkFrame(app)
        self.options.grid(row=0, column=0, sticky="nsew", rowspan=10)
        self.options.bind_all("<Escape>", lambda event: self.destroy_window())
        self.drm_label = ctk.CTkLabel(self.options, text="Activation-Key:")
        self.drm_label.grid(row=0, column=0, sticky="ew", padx=10)
        self.drm_entry = ctk.CTkEntry(self.options, placeholder_text_color="white")
        self.drm_entry.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        self.drm_entry.insert(0, self.drm_key)
        self.drm_entry.configure(state="disabled", justify="center")
        self.edit_button = ctk.CTkButton(self.options, text="Edit", command=self.update_drm, width=230)
        self.edit_button.grid(row=2, column=0, pady=10, padx=10, sticky="w")
        # self.save_button = ctk.CTkButton(self.options, text="Save", command=self.update_drm, width=110)
        # self.save_button.grid(row=2, column=0, pady=10, padx=10, sticky="e")
        # TODO: Spec: Add editbox with pixel. Disable until edit button pressed
        self.specs_combo = ctk.CTkComboBox(self.options, values=self.get_specs(), width=110, command=self.get_spells)
        self.specs_combo.grid(row=4, column=0, sticky="ew", padx=10, pady=10)
        self.specs_combo.set("Select spec")
        # TODO: Create add spec function
        self.specs_add = ctk.CTkButton(self.options, text="Add", command=lambda: call_spec(),
                                       width=70)
        self.specs_add.grid(row=5, column=0, pady=10, padx=10, sticky="w")
        # TODO: Create edit spell function and rename button to "edit"
        self.specs_rename = ctk.CTkButton(self.options, text="Rename",
                                          command=lambda: self.rename_spec(self.specs_combo.get()), width=70)
        self.specs_rename.grid(row=5, column=0, pady=10, padx=10, sticky="n")
        self.specs_del = ctk.CTkButton(self.options, text="Delete",
                                       command=lambda: confirmation_box("Delete Spec?",
                                                                        f"Are you sure you want to delete {self.specs_combo.get()}?",
                                                                        self.del_spec), width=70)

        # TODO: Spells: Add editbox with pixel and hotkey. Disable until edit button pressed
        self.specs_del.grid(row=5, column=0, pady=10, padx=10, sticky="e")

        self.spells_combo = ctk.CTkComboBox(self.options, width=110)
        self.spells_combo.set("Select spell")
        self.spells_combo.configure(state="disabled")
        self.spells_combo.grid(row=6, column=0, sticky="ew", padx=10, pady=10)
        # TODO: Create add spell function
        self.spells_add = ctk.CTkButton(self.options, text="Add", command=lambda: call_spell(), width=70)
        self.spells_add.grid(row=7, column=0, pady=10, padx=10, sticky="w")
        # TODO: Create edit spell function and rename button to "edit"
        self.spells_rename = ctk.CTkButton(self.options, text="Rename",
                                           command=lambda: self.rename_spell(self.spells_combo.get()), width=70)
        self.spells_rename.grid(row=7, column=0, pady=10, padx=10, sticky="n")
        self.spells_del = ctk.CTkButton(self.options, text="Delete",
                                        command=lambda: confirmation_box("Delete Spell?",
                                                                         f"Are you sure you want to delete {self.spells_combo.get()}?",
                                                                         self.del_spell), width=70)
        self.spells_del.grid(row=7, column=0, pady=10, padx=10, sticky="e")

        self.close_button = ctk.CTkButton(self.options, text="Close", command=self.destroy_window, width=230)
        self.close_button.grid(row=8, column=0, pady=10, padx=10)

    def update_drm(self) -> None:
        self.drm_entry.get()
        if self.drm_entry.state == "disabled":
            self.drm_entry.configure(state="normal")
            self.drm_entry.focus()
            self.edit_button.configure(text="Save")
        else:
            self.drm_entry.configure(state="disabled")
            self.edit_button.configure(text="Edit")
            config = get_config()
            config["drm"] = self.drm_entry.get()
            save_config(config)

    def destroy_window(self):
        self.drm_entry.delete(0)
        self.options.destroy()

    def get_specs(self):
        # config = get_config()
        for profile in self.config['profiles']:
            if profile['name'] == self.config['current_profile']:
                return [spec["name"] for spec in profile['specs']]

    def get_spec(self):
        for spec in self.get_specs():
            if spec == self.specs_combo.get():
                return spec

    def update_specs(self):
        # config = get_config()
        for profile in self.config['profiles']:
            self.specs_combo.configure(values=[spec["name"] for spec in profile['specs']])
            if rotation.active_window():
                self.specs_combo.set(self.get_spec())
            else:
                self.specs_combo.set(profile["specs"][0]["name"])
            self.get_spells(profile["specs"][0]["name"])

    def add_spec(self):
        pass

    def del_spec(self):
        config = get_config()
        for profile in enumerate(get_profiles()):
            if profile[1] == profiles.get():
                for spec in enumerate(config["profiles"][profile[0]]["specs"]):
                    if spec[1]['name'] == self.specs_combo.get():
                        config["profiles"][profile[0]]["specs"].pop(spec[0])
                        save_config(config)
                        self.update_specs()

    def rename_spec(self):
        pass
    def get_spells(self, value):
        config = get_config()
        for profile in config['profiles']:
            if profile['name'] == config['current_profile']:
                for spec in enumerate(profile['specs']):
                    if spec[1]['name'] == value:
                        try:
                            self.spells_combo.configure(state="normal")
                            unsorted_list = [spell for spell in spec[1]['spells']]
                            for spell in unsorted_list:
                                if spell["key"] == "{VK_NUMPAD0}":
                                    unsorted_list.remove(spell)
                            sorted_list = sorted(spell["name"] for spell in unsorted_list)
                            self.spells_combo.configure(values=sorted_list)
                            self.spells_combo.set(sorted_list[0])
                            return
                        except IndexError:
                            self.spells_combo.configure(state="disabled")
                            return
                else:
                    self.spells_combo.configure(state="disabled")
                    self.spells_combo.configure(values=[])
                    self.spells_combo.set("Select spell")

    def del_spell(self):
        config = get_config()
        for profile in enumerate(get_profiles()):
            if profile[1] == profiles.get():
                for spec in enumerate(config["profiles"][profile[0]]["specs"]):
                    if spec[1]['name'] == self.specs_combo.get():
                        for spell in enumerate(spec[1]['spells']):
                            if spell[1]['name'] == self.spells_combo.get():
                                spec[1]['spells'].pop(spell[0])
                                save_config(config)
                                self.get_spells(self.specs_combo.get())

    def rename_spell(self):
        pass


class Spell:
    def __init__(self):
        super().__init__()
        self.new_spell = ctk.CTkToplevel(app)
        self.new_spell.title("Add Spell")
        self.hotkey = []
        self.last_key = []
        self.vk = []
        self.vk_key = []
        try:
            self.pixel_value = get_spell()[2]
        except IndexError:
            self.pixel_value = "No Spec detected"
        self.spec = get_spec()[0]
        self.label = ctk.CTkLabel(self.new_spell, text="Enter spell").grid(row=0, column=0, columnspan=2)
        self.key = ctk.CTkEntry(self.new_spell)
        self.key.grid(row=2, column=0, columnspan=2, pady=10, padx=5)
        self.key.configure(placeholder_text="Hotkey")
        self.key.bind("<Key>", self.get_hotkey)
        self.key.bind("<KeyRelease>", self.get_hotkey_release)
        self.key.bind("<Escape>", lambda event: self.key.delete(0, "end"))
        self.pixel = ctk.CTkEntry(self.new_spell)
        self.pixel.grid(row=3, column=0, columnspan=2, pady=10, padx=5)
        try:
            self.pixel.insert(0, self.pixel_value)
        except IndexError:
            self.pixel.insert(0, "Could not detect spec. Is WoW in focus?")
        self.pixel.configure(state="disabled")
        self.ok_button = ctk.CTkButton(self.new_spell, command=lambda: [self.add_spells(), self.new_spell.destroy()],
                                       text="Ok", width=75)
        self.ok_button.grid(row=4, column=0, pady=10, padx=5)
        self.cancel_button = ctk.CTkButton(self.new_spell, text="Cancel", command=lambda: self.new_spell.destroy(),
                                           width=75)
        self.cancel_button.grid(row=4, column=1, pady=10, padx=5)
        self.new_spell.bind("<Escape>", lambda event: [self.key.configure(state="normal"), self.key.delete(0, "end"),
                                                       self.key.focus()])
        self.name = ctk.CTkEntry(self.new_spell)
        self.name.grid(row=1, column=0, columnspan=2, pady=10, padx=5)
        self.name.configure(placeholder_text="Spell name")
        self.name.bind("<Return>", lambda event: self.key.focus_force())
        self.name.bind("<Escape>", lambda event: self.name.delete(0, "end"))
        self.key.bind("<Return>", lambda event: [self.add_spells(), self.new_spell.destroy()])
        self.new_spell.bind("<Return>", lambda event: [self.add_spells(), self.new_spell.destroy()])

    def add_spells(self):
        try:
            print(self.name.get(), self.cfg_hotkey, self.pixel.get(), self.spec)
            add_spell(self.name.get(), self.cfg_hotkey, self.pixel_value, self.spec)
        except IndexError:
            add_spell(self.name.get(), self.cfg_hotkey, self.pixel_value, "None")  # TODO: Add spec detection

    def get_hotkey(self, event):
        self.vk = []
        if event.keycode == self.last_key:
            return
        elif event.keycode == 27:
            self.hotkey = []
            self.vk = []
            return
        else:
            for key, value in vkc.items():
                if key == event.keycode:
                    if len(value) > 1:
                        value = value.title()
                    if len(value) == 1:
                        value = value
                    self.hotkey.append(value)
            for key in self.hotkey:
                if len(key) > 1:
                    self.vk.append(key.upper())
                else:
                    self.vk.append(key)

        self.cfg_hotkey = "".join(self.vk)  # Use this to save to config
        self.last_key = event.keycode
        self.last_key = event.keycode
        return

    def get_hotkey_release(self, event):
        self.key.delete(0, "end")
        for keys in enumerate(self.hotkey):
            if keys[1] == "^":
                self.hotkey[keys[0]] = "Ctrl"
            elif keys[1] == "+":
                self.hotkey[keys[0]] = "Shift"
            elif keys[1] == "%":
                self.hotkey[keys[0]] = "Alt"
        self.key.insert(0, " + ".join(self.hotkey))
        if event.keycode not in [9, 13, 27]:
            self.key.configure(state="disabled")
            self.new_spell.focus()
        if event.keycode == 27:
            self.key.configure(state="normal")
            self.key.delete(0, "end")
            self.hotkey = []
            self.vk = []


class Spec:
    def __init__(self):
        super().__init__()
        self.new_spec = ctk.CTkToplevel(app)
        self.new_spec.title("Add Spec")
        self.pixel_value = get_spec()[2]
        self.label = ctk.CTkLabel(self.new_spec, text="Enter Spec Name:").grid(row=0, column=0, columnspan=2)
        self.pixel = ctk.CTkEntry(self.new_spec)
        self.pixel.grid(row=3, column=0, columnspan=2, pady=10, padx=5)
        self.pixel.insert(0, self.pixel_value)
        self.pixel.configure(state="disabled")
        self.ok_button = ctk.CTkButton(self.new_spec, text="Ok",
                                       command=lambda: [self.add_specs, self.new_spec.destroy()], width=75)
        self.ok_button.grid(row=4, column=0, pady=10, padx=5)
        self.cancel_button = ctk.CTkButton(self.new_spec, text="Cancel", command=lambda: self.new_spec.destroy(),
                                           width=75)
        self.cancel_button.grid(row=4, column=1, pady=10, padx=5)
        self.name = ctk.CTkEntry(self.new_spec)
        self.name.grid(row=1, column=0, columnspan=2, pady=10, padx=5)
        self.name.configure(placeholder_text="Spec Name")
        self.name.bind("<Return>", lambda event: [self.add_specs(), self.new_spec.destroy()])
        self.new_spec.bind("<Return>", lambda event: [self.add_specs(), self.new_spec.destroy()])

    def add_specs(self):
        try:
            print(self.name.get(), self.pixel_value)
            add_spec(self.name.get(), self.pixel_value)
        except IndexError:
            add_spec(self.name.get(), self.pixel_value)  # TODO: Add spec detection
        self.new_spec.destroy()

# Use CTkButton instead of tkinter Button
# TODO: Add editbutton and allow editing of all profile related settings
profiles = ctk.CTkOptionMenu(app, values=update_profiles(), command=set_profile, width=230)
profiles.grid(row=0, column=0, pady=10, padx=10)

# TODO: Make buttons smaller and place them next to each other
button_new = ctk.CTkButton(app, text="New Profile", border_color="black", command=add_profile, width=230)
button_new.grid(row=1, column=0, pady=10, padx=10)

button_del = ctk.CTkButton(app, text="Delete Profile",
                           width=230, command=lambda: confirmation_box("Delete Profile",
                                                                       "Are you sure you want to delete this profile?",
                                                                       del_profile))
button_del.grid(row=2, column=0, pady=10, padx=10)
# button_del.bind("<Button-1>", lambda event: delete_profile(optionmenu_1.get()))

button_config = ctk.CTkButton(app, text="Open Config", command=lambda: webbrowser.open("config.json"), width=230)
button_config.grid(row=3, column=0, pady=10, padx=10)

button_ahk = ctk.CTkButton(app, text="Import AHK", command=import_ahk, width=230)
button_ahk.grid(row=4, column=0, pady=10, padx=10)

button_options = ctk.CTkButton(app, text="More Options", command=OptionsMenu, width=230)
button_options.grid(row=5, column=0, pady=10, padx=10)

slider = ctk.CTkSlider(app, from_=10, to=100, number_of_steps=90, command=repeat_delay,
                       width=230)
slider.grid(row=7, column=0, pady=10, padx=10)
slider.set(get_repeat_delay() * 1000)

delay_label = ctk.CTkLabel(app, text=f"Delay between spells (ms): {int(slider.get())}")
delay_label.grid(row=6, column=0, pady=5, padx=10)

toggle = ctk.IntVar()
switch_1 = ctk.CTkSwitch(app, text=f"AFK Mode: Off", variable=toggle, onvalue=1, offvalue=0)
switch_1.grid(row=8, column=0, pady=10, padx=10)
switch_1.configure(command=update_toggle)


# add_spell_hotkey = ctk.CTkButton(app)
# add_spell_hotkey.grid(row=0, column=1)
# add_spell_hotkey.bind_all("<Key>", lambda: [get_hotkey, print(get_hotkey)])
async def call_spec():
    spec = Spec()
    await asyncio.sleep(0.2)
    spec.new_spec.focus()
    spec.name.focus()


async def call_spell():
    spell = Spell()
    await asyncio.sleep(0.2)
    spell.new_spell.focus()
    spell.name.focus()


async def hotkey_listener():
    while True:
        name = get_config()["current_profile"]
        modifier = None
        for profile in get_config()["profiles"]:
            if profile["name"] == name:
                spec_hotkey = profile["spec_hotkey"]
                spell_hotkey = profile["spell_hotkey"]
                if spec_hotkey[0] in ["+", "^", "%"]:
                    modifier = spec_hotkey[0]
                    spec_hotkey = spec_hotkey[1:]
                if spell_hotkey[0] in ["+", "^", "%"]:
                    modifier = spell_hotkey[0]
                    spell_hotkey = spell_hotkey[1:]
                for k, v in vkc.items():
                    if v == spec_hotkey:
                        spec_hotkey = k
                    if v == spell_hotkey:
                        spell_hotkey = k
                    if v == modifier:
                        modifier = k
                hotkey = spec_hotkey, spell_hotkey, modifier
        try:
            if hotkey[2]:
                if win32api.GetAsyncKeyState(hotkey[0]) < 0 and win32api.GetAsyncKeyState(hotkey[2]) < 0:
                    await call_spec()
                if win32api.GetAsyncKeyState(hotkey[1]) < 0 and win32api.GetAsyncKeyState(hotkey[2]) < 0:
                    await call_spell()
            else:
                if win32api.GetAsyncKeyState(hotkey[0]) < 0:
                    await call_spec()
                if win32api.GetAsyncKeyState(hotkey[1]) < 0:
                    await call_spell()
            await asyncio.sleep(0.1)
        except KeyError:
            pass
        except UnboundLocalError:
            pass
        except TypeError:
            pass
        await asyncio.sleep(0.01)


async def run_tk(root, interval=1 / 60):
    try:
        while root.state():
            root.update()
            await asyncio.sleep(interval)
    except TclError as e:
        if "application has been destroyed" not in e.args[0]:
            raise

async def shift_state():
    while True:
        if win32api.GetAsyncKeyState(0x10) < 0:
            log.info(win32api.GetAsyncKeyState(16))
        await asyncio.sleep(0.00001)


async def main():
    asyncio.create_task(rotation.main())
    asyncio.create_task(hotkey_listener())
    # asyncio.create_task(shift_state())
    await asyncio.create_task(run_tk(app))
    sys.exit()


if __name__ == "__main__":
    loop = asyncio.get_event_loop()
    loop.run_until_complete(main())
    loop.close()
