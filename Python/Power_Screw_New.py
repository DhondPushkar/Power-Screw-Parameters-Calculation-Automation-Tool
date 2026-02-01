import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import math

# ---------------- CALCULATION LOGIC ---------------- #

def calculate_power_screw(row):
    pitch = row['Pitch']
    threads = row['Threads']
    nominal_dia = row['NominalDiameter']
    core_dia = row['CoreDiameter']
    friction_angle = math.radians(row['FrictionAngle'])
    load = row['Load']

    mean_dia = nominal_dia - 0.5 * pitch
    lead = threads * pitch
    helix_angle = math.atan(lead / (math.pi * mean_dia))

    torque_raise = (load * mean_dia * math.tan(friction_angle + helix_angle)) / 2
    torque_lower = (load * mean_dia * math.tan(friction_angle - helix_angle)) / 2

    efficiency = math.tan(helix_angle) / math.tan(helix_angle + friction_angle)
    max_efficiency = (1 - math.sin(friction_angle)) / (1 + math.sin(friction_angle))
    overall_efficiency = (load * lead) / (2 * math.pi)

    shear_stress = (16 * torque_raise) / (math.pi * core_dia ** 3)
    compressive_stress = (4 * load) / (math.pi * core_dia ** 2)

    max_principal_stress = (
        (-compressive_stress / 2) +
        math.sqrt((compressive_stress / 2) ** 2 + shear_stress ** 2)
    )

    max_shearing_stress = math.sqrt(
        compressive_stress ** 2 + shear_stress ** 2
    )

    return pd.Series({
        "MeanDiameter": round(mean_dia, 3),
        "Lead": round(lead, 3),
        "HelixAngle(deg)": round(math.degrees(helix_angle), 3),
        "TorqueRaise": round(torque_raise, 3),
        "TorqueLower": round(torque_lower, 3),
        "Efficiency": round(efficiency, 3),
        "MaximumEfficiency": round(max_efficiency, 3),
        "OverallEfficiency": round(overall_efficiency, 3),
        "ShearStress": round(shear_stress, 3),
        "CompressiveStress": round(compressive_stress, 3),
        "MaxPrincipalStress": round(max_principal_stress, 3),
        "MaxShearingStress": round(max_shearing_stress, 3)
    })

# ---------------- EXCEL HANDLING ---------------- #

def upload_and_calculate():
    try:
        input_file = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not input_file:
            return

        df = pd.read_excel(input_file)

        required_cols = [
            'Pitch', 'Threads', 'NominalDiameter',
            'CoreDiameter', 'FrictionAngle', 'Load'
        ]

        if not all(col in df.columns for col in required_cols):
            raise ValueError("Excel file missing required columns")

        result_df = df.join(df.apply(calculate_power_screw, axis=1))

        save_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx"
        )
        result_df.to_excel(save_file, index=False)

        messagebox.showinfo(
            "Success",
            "All power screw parameters calculated and saved successfully."
        )

    except Exception as e:
        messagebox.showerror("Error", str(e))

# ---------------- UI ---------------- #

root = tk.Tk()
root.title("Power Screw Performance Automation Tool")
root.geometry("500x320")

title = ttk.Label(
    root,
    text="Power Screw Calculation Tool",
    font=("Arial", 14, "bold")
)
title.pack(pady=15)

instruction = ttk.Label(
    root,
    text=(
        "Upload Excel with columns:\n"
        "Pitch, Threads, NominalDiameter,\n"
        "CoreDiameter, FrictionAngle, Load"
    ),
    justify="center"
)
instruction.pack(pady=10)

btn = ttk.Button(
    root,
    text="Upload Excel & Calculate",
    command=upload_and_calculate
)
btn.pack(pady=20)

root.mainloop()
