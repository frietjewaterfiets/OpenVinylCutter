import io
import math
import re
import subprocess
import sys
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import serial
import win32print
from serial.tools import list_ports
from svgelements import Path as SvgPath
from svgelements import SVG, Shape


MM_PER_INCH = 25.4
PX_PER_INCH = 96.0
MM_PER_PX = MM_PER_INCH / PX_PER_INCH


class SvgPlotterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("OpenVinylCutter")
        self.root.geometry("1180x760")
        self.root.minsize(980, 620)
        self.app_dir = Path(__file__).resolve().parent
        self.header_logo_path = self.app_dir / "ovc-logo.png"

        self.svg_file: Path | None = None
        self.polylines_mm: list[list[tuple[float, float]]] = []
        self.bounds_mm: tuple[float, float, float, float] | None = None
        self.header_logo_image: tk.PhotoImage | None = None

        self._build_variables()
        self._build_layout()
        self.refresh_connections()

    def _build_variables(self) -> None:
        self.transport_var = tk.StringVar(value="Automatic")
        self.port_var = tk.StringVar()
        self.printer_var = tk.StringVar()
        self.usb_port_var = tk.StringVar()
        self.baud_var = tk.StringVar(value="9600")
        self.width_var = tk.StringVar(value="580")
        self.height_var = tk.StringVar(value="1000")
        self.margin_var = tk.StringVar(value="2")
        self.units_var = tk.StringVar(value="40")
        self.status_var = tk.StringVar(value="Ready. Load / drag & drop an SVG to begin.")
        self.file_var = tk.StringVar(value="No SVG loaded")
        self.language_var = tk.StringVar(value="HPGL")
        self.mirror_var = tk.BooleanVar(value=False)
        self.rotate_var = tk.BooleanVar(value=False)
        self.fit_width_var = tk.BooleanVar(value=False)
        self.peel_box_var = tk.BooleanVar(value=True)
        self.flip_x_var = tk.BooleanVar(value=False)
        self.flip_y_var = tk.BooleanVar(value=True)
        self.overcut_var = tk.StringVar(value="0")
        self.feed_after_var = tk.StringVar(value="40")
        self.feed_axis_var = tk.StringVar(value="X")
        self.feed_direction_var = tk.StringVar(value="+")

    def _build_layout(self) -> None:
        self.root.columnconfigure(0, weight=0)
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=0)

        controls_wrap = ttk.Frame(self.root, padding=(16, 16, 8, 8))
        controls_wrap.grid(row=0, column=0, sticky="nsw")
        preview = ttk.Frame(self.root, padding=(0, 16, 16, 8))
        preview.grid(row=0, column=1, sticky="nsew")
        controls_wrap.rowconfigure(0, weight=1)
        controls_wrap.columnconfigure(0, weight=1)

        controls_canvas = tk.Canvas(
            controls_wrap,
            highlightthickness=0,
            borderwidth=0,
            width=360,
        )
        controls_scrollbar = ttk.Scrollbar(
            controls_wrap, orient="vertical", command=controls_canvas.yview
        )
        controls = ttk.Frame(controls_canvas, padding=(0, 0, 12, 0))
        controls.bind(
            "<Configure>",
            lambda _event: controls_canvas.configure(
                scrollregion=controls_canvas.bbox("all")
            ),
        )
        controls_canvas.create_window((0, 0), window=controls, anchor="nw")
        controls_canvas.configure(yscrollcommand=controls_scrollbar.set)
        controls_canvas.grid(row=0, column=0, sticky="nsw")
        controls_scrollbar.grid(row=0, column=1, sticky="ns")

        for idx in range(10):
            controls.rowconfigure(idx, weight=0)
        controls.rowconfigure(10, weight=1)

        if self.header_logo_path.exists():
            try:
                self.header_logo_image = tk.PhotoImage(file=str(self.header_logo_path))
                logo_width = max(self.header_logo_image.width(), 1)
                max_logo_width = 370
                if logo_width > max_logo_width:
                    shrink = max(1, math.ceil(logo_width / max_logo_width))
                    self.header_logo_image = self.header_logo_image.subsample(shrink, shrink)
            except tk.TclError:
                self.header_logo_image = None

        if self.header_logo_image is not None:
            header = ttk.Label(controls, image=self.header_logo_image)
        else:
            header = ttk.Label(
                controls,
                text="OpenVinylCutter",
                font=("Segoe UI Semibold", 18),
            )
        header.grid(row=0, column=0, sticky="w", pady=(0, 14))
        byline = tk.Label(
            controls,
            text="by frietjewaterfiets.",
            fg="#2f81f7",
            cursor="hand2",
            font=("Segoe UI", 10),
            anchor="w",
            justify="left",
        )
        byline.grid(row=1, column=0, sticky="w", pady=(0, 12))
        byline.bind(
            "<Button-1>",
            lambda _event: webbrowser.open("https://github.com/frietjewaterfiets"),
        )

        ttk.Button(controls, text="Load SVG", command=self.load_svg).grid(
            row=2, column=0, sticky="ew"
        )
        ttk.Label(
            controls, textvariable=self.file_var, wraplength=320, justify="left"
        ).grid(row=3, column=0, sticky="ew", pady=(8, 14))

        device_frame = ttk.LabelFrame(controls, text="Plotter Connection", padding=12)
        device_frame.grid(row=4, column=0, sticky="ew", pady=(0, 12))
        device_frame.columnconfigure(1, weight=1)

        ttk.Label(device_frame, text="Connection Type").grid(row=0, column=0, sticky="w")
        ttk.Combobox(
            device_frame,
            textvariable=self.transport_var,
            state="readonly",
            values=["Automatic", "Serial (COM)", "Windows USB/Printer"],
        ).grid(row=0, column=1, sticky="ew", padx=(8, 0))

        ttk.Label(device_frame, text="COM port").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.port_combo = ttk.Combobox(
            device_frame, textvariable=self.port_var, state="readonly", width=24
        )
        self.port_combo.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(10, 0))

        ttk.Label(device_frame, text="USB printer queue").grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.printer_combo = ttk.Combobox(
            device_frame, textvariable=self.printer_var, state="readonly", width=24
        )
        self.printer_combo.grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=(10, 0))

        ttk.Label(device_frame, text="USB port").grid(row=3, column=0, sticky="w", pady=(10, 0))
        self.usb_port_combo = ttk.Combobox(
            device_frame, textvariable=self.usb_port_var, state="readonly", width=24
        )
        self.usb_port_combo.grid(row=3, column=1, sticky="ew", padx=(8, 0), pady=(10, 0))

        button_row = ttk.Frame(device_frame)
        button_row.grid(row=4, column=1, sticky="ew", pady=(8, 0))
        button_row.columnconfigure(0, weight=1)
        button_row.columnconfigure(1, weight=1)
        ttk.Button(button_row, text="Refresh", command=self.refresh_connections).grid(
            row=0, column=0, sticky="ew", padx=(0, 4)
        )
        ttk.Button(button_row, text="Create USB Queue", command=self.create_usb_queue).grid(
            row=0, column=1, sticky="ew", padx=(4, 0)
        )

        ttk.Label(device_frame, text="Baud rate").grid(row=5, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(device_frame, textvariable=self.baud_var).grid(
            row=5, column=1, sticky="ew", padx=(8, 0), pady=(12, 0)
        )

        setup_frame = ttk.LabelFrame(controls, text="Cut Settings", padding=12)
        setup_frame.grid(row=5, column=0, sticky="ew", pady=(0, 12))
        setup_frame.columnconfigure(1, weight=1)

        ttk.Label(setup_frame, text="Width (mm)").grid(row=0, column=0, sticky="w")
        ttk.Entry(setup_frame, textvariable=self.width_var).grid(
            row=0, column=1, sticky="ew", padx=(8, 0)
        )
        ttk.Label(setup_frame, text="Height (mm)").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(setup_frame, textvariable=self.height_var).grid(
            row=1, column=1, sticky="ew", padx=(8, 0), pady=(8, 0)
        )
        ttk.Label(setup_frame, text="Peel box offset (mm)").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(setup_frame, textvariable=self.margin_var).grid(
            row=2, column=1, sticky="ew", padx=(8, 0), pady=(8, 0)
        )
        ttk.Label(setup_frame, text="HPGL units/mm").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(setup_frame, textvariable=self.units_var).grid(
            row=3, column=1, sticky="ew", padx=(8, 0), pady=(8, 0)
        )
        ttk.Label(setup_frame, text="Overcut (mm)").grid(row=4, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(setup_frame, textvariable=self.overcut_var).grid(
            row=4, column=1, sticky="ew", padx=(8, 0), pady=(8, 0)
        )
        ttk.Label(setup_frame, text="Feed after cut (mm)").grid(row=5, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(setup_frame, textvariable=self.feed_after_var).grid(
            row=5, column=1, sticky="ew", padx=(8, 0), pady=(8, 0)
        )
        ttk.Label(setup_frame, text="Feed axis").grid(row=6, column=0, sticky="w", pady=(8, 0))
        ttk.Combobox(
            setup_frame,
            textvariable=self.feed_axis_var,
            state="readonly",
            values=["X", "Y"],
        ).grid(row=6, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))
        ttk.Label(setup_frame, text="Feed direction").grid(row=7, column=0, sticky="w", pady=(8, 0))
        ttk.Combobox(
            setup_frame,
            textvariable=self.feed_direction_var,
            state="readonly",
            values=["+", "-"],
        ).grid(row=7, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        options_frame = ttk.LabelFrame(controls, text="Options", padding=12)
        options_frame.grid(row=6, column=0, sticky="ew", pady=(0, 12))
        options_frame.columnconfigure(0, weight=1)
        ttk.Checkbutton(
            options_frame, text="Mirror horizontally", variable=self.mirror_var, command=self.redraw_preview
        ).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(
            options_frame, text="Rotate 90 degrees", variable=self.rotate_var, command=self.redraw_preview
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Checkbutton(
            options_frame,
            text="Fit to material width automatically (off = 1:1)",
            variable=self.fit_width_var,
            command=self.redraw_preview,
        ).grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Checkbutton(
            options_frame,
            text="Cut peel box",
            variable=self.peel_box_var,
            command=self.redraw_preview,
        ).grid(row=3, column=0, sticky="w", pady=(6, 0))
        ttk.Checkbutton(
            options_frame,
            text="Flip plotter X axis",
            variable=self.flip_x_var,
            command=self.redraw_preview,
        ).grid(row=4, column=0, sticky="w", pady=(6, 0))
        ttk.Checkbutton(
            options_frame,
            text="Flip plotter Y axis",
            variable=self.flip_y_var,
            command=self.redraw_preview,
        ).grid(row=5, column=0, sticky="w", pady=(6, 0))

        ttk.Label(options_frame, text="Language").grid(row=6, column=0, sticky="w", pady=(10, 0))
        ttk.Combobox(
            options_frame,
            textvariable=self.language_var,
            state="readonly",
            values=["HPGL"],
        ).grid(row=7, column=0, sticky="ew", pady=(6, 0))

        action_frame = ttk.Frame(controls)
        action_frame.grid(row=7, column=0, sticky="ew")
        action_frame.columnconfigure(0, weight=1)
        action_frame.columnconfigure(1, weight=1)

        ttk.Button(action_frame, text="Refresh Preview", command=self.redraw_preview).grid(
            row=0, column=0, sticky="ew", padx=(0, 6)
        )
        ttk.Button(action_frame, text="Save HPGL", command=self.save_hpgl).grid(
            row=0, column=1, sticky="ew", padx=(6, 0)
        )

        ttk.Button(
            controls,
            text="Send to Plotter",
            command=self.send_to_plotter,
        ).grid(row=8, column=0, sticky="ew", pady=(12, 0))

        hint = (
            "Tip: if the cut comes out too large or too small, adjust "
            "'HPGL units/mm' first. 40 is a good starting point."
        )
        ttk.Label(controls, text=hint, wraplength=320, justify="left").grid(
            row=9, column=0, sticky="ew", pady=(14, 0)
        )

        controls_canvas.bind_all(
            "<MouseWheel>",
            lambda event: controls_canvas.yview_scroll(int(-event.delta / 120), "units"),
        )

        preview.columnconfigure(0, weight=1)
        preview.rowconfigure(0, weight=1)
        self.canvas = tk.Canvas(preview, background="#10151b", highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.canvas.bind("<Configure>", lambda _event: self.redraw_preview())

        status = ttk.Label(
            self.root, textvariable=self.status_var, anchor="w", padding=(16, 8)
        )
        status.grid(row=1, column=0, columnspan=2, sticky="ew")

    def refresh_connections(self) -> None:
        ports = []
        for port in list_ports.comports():
            label = f"{port.device} - {port.description}"
            ports.append(label)

        self.port_combo["values"] = ports
        if ports and (not self.port_var.get() or self.port_var.get() not in ports):
            self.port_var.set(ports[0])
        elif not ports:
            self.port_var.set("")
        printers = self._list_printers()
        self.printer_combo["values"] = printers
        if printers and (not self.printer_var.get() or self.printer_var.get() not in printers):
            self.printer_var.set(printers[0])
        elif not printers:
            self.printer_var.set("")

        usb_ports = self._list_usb_ports()
        self.usb_port_combo["values"] = usb_ports
        if usb_ports and (not self.usb_port_var.get() or self.usb_port_var.get() not in usb_ports):
            self.usb_port_var.set(usb_ports[0])
        elif not usb_ports:
            self.usb_port_var.set("")

        self.status_var.set(
            f"Found {len(ports)} COM ports, {len(printers)} printer queues, and {len(usb_ports)} USB ports."
        )

    def _list_printers(self) -> list[str]:
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        queues = []
        for info in win32print.EnumPrinters(flags, None, 2):
            name = info["pPrinterName"] or ""
            port_name = info["pPortName"] or ""
            driver_name = info["pDriverName"] or ""
            if self._is_usb_port(port_name) or "plotter" in name.lower():
                queues.append(f"{name} [{port_name}] - {driver_name}")
        return queues

    def _list_usb_ports(self) -> list[str]:
        ports = []
        try:
            for info in win32print.EnumPorts(None, 2):
                port_name = info.get("Name", "")
                if self._is_usb_port(port_name):
                    ports.append(port_name)
        except Exception:
            return []
        return sorted(set(ports))

    def _is_usb_port(self, port_name: str) -> bool:
        return bool(re.fullmatch(r"USB\d{3}", (port_name or "").strip(), flags=re.IGNORECASE))

    def create_usb_queue(self) -> None:
        usb_port = self.usb_port_var.get().strip()
        if not usb_port:
            messagebox.showinfo("No USB Port", "Choose a USB port first.")
            return

        queue_name = f"OpenVinylCutter RAW {usb_port}"
        existing = [entry for entry in self._list_printers() if entry.startswith(queue_name + " [")]
        if existing:
            self.printer_var.set(existing[0])
            self.transport_var.set("Windows USB/Printer")
            self.status_var.set(f"USB queue already exists: {queue_name}")
            return

        command = (
            f"Add-Printer -Name '{queue_name}' "
            f"-DriverName 'Microsoft enhanced Point and Print compatibility driver' "
            f"-PortName '{usb_port}'"
        )
        try:
            subprocess.run(
                ["powershell", "-NoProfile", "-Command", command],
                check=True,
                capture_output=True,
                text=True,
            )
        except subprocess.CalledProcessError as exc:
            details = (exc.stderr or exc.stdout or str(exc)).strip()
            messagebox.showerror(
                "USB Queue Failed",
                "Could not create a Windows printer queue.\n\n" + details,
            )
            return

        self.refresh_connections()
        for entry in self._list_printers():
            if entry.startswith(queue_name + " ["):
                self.printer_var.set(entry)
                break
        self.transport_var.set("Windows USB/Printer")
        self.status_var.set(f"USB queue created on {usb_port}: {queue_name}")

    def load_svg(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Choose an SVG",
            filetypes=[("SVG files", "*.svg")],
        )
        if not file_path:
            return

        try:
            self._load_svg_path(Path(file_path))
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("SVG Error", f"Could not load SVG:\n{exc}")
            return

    def _load_svg_path(self, svg_path: Path) -> None:
        polylines, bounds = self._parse_svg(svg_path)
        self.svg_file = svg_path
        self.polylines_mm = polylines
        self.bounds_mm = bounds
        self.file_var.set(str(self.svg_file))
        self.status_var.set(
            f"SVG loaded: {len(self.polylines_mm)} cut paths found."
        )
        self.redraw_preview()

    def open_svg_on_startup(self, svg_path: Path) -> None:
        try:
            self._load_svg_path(svg_path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Startup SVG Error", f"Could not open startup SVG:\n{exc}")

    def _parse_svg(
        self, svg_path: Path
    ) -> tuple[list[list[tuple[float, float]]], tuple[float, float, float, float]]:
        data = svg_path.read_text(encoding="utf-8", errors="ignore")
        svg = SVG.parse(io.StringIO(data), reify=True, ppi=PX_PER_INCH)

        polylines: list[list[tuple[float, float]]] = []
        seen_signatures: set[tuple[tuple[int, int], ...]] = set()
        min_x = math.inf
        min_y = math.inf
        max_x = -math.inf
        max_y = -math.inf

        for element in svg.elements():
            if not isinstance(element, Shape):
                continue
            values = getattr(element, "values", {})
            if values.get("visibility") == "hidden":
                continue
            if values.get("display") == "none":
                continue
            if str(values.get("opacity", "1")).strip() == "0":
                continue
            if str(values.get("fill", "")).strip().lower() == "none" and str(
                values.get("stroke", "")
            ).strip().lower() == "none":
                continue

            path = SvgPath(element)
            for subpath in path.as_subpaths():
                polyline = self._subpath_to_polyline(subpath)
                if len(polyline) < 2:
                    continue
                polyline_mm = []
                for point in polyline:
                    x_mm = float(point.real) * MM_PER_PX
                    y_mm = float(point.imag) * MM_PER_PX
                    polyline_mm.append((x_mm, y_mm))
                cleaned = self._dedupe_points(polyline_mm)
                if len(cleaned) < 2:
                    continue
                signature = self._polyline_signature(cleaned)
                if signature in seen_signatures:
                    continue
                seen_signatures.add(signature)
                for x_mm, y_mm in cleaned:
                    min_x = min(min_x, x_mm)
                    min_y = min(min_y, y_mm)
                    max_x = max(max_x, x_mm)
                    max_y = max(max_y, y_mm)
                polylines.append(cleaned)

        if not polylines:
            raise ValueError("No plottable paths were found in this SVG.")

        return polylines, (min_x, min_y, max_x, max_y)

    def _subpath_to_polyline(self, subpath) -> list[complex]:
        points: list[complex] = []
        for segment in subpath:
            length = max(float(segment.length(error=1e-3)), 0.1)
            steps = max(2, min(250, int(length / 3)))
            for idx in range(steps + 1):
                t = idx / steps
                points.append(segment.point(t))
        return points

    def _dedupe_points(
        self, points: list[tuple[float, float]], tolerance: float = 0.02
    ) -> list[tuple[float, float]]:
        cleaned: list[tuple[float, float]] = []
        for point in points:
            if not cleaned:
                cleaned.append(point)
                continue
            last = cleaned[-1]
            if abs(last[0] - point[0]) > tolerance or abs(last[1] - point[1]) > tolerance:
                cleaned.append(point)
        return cleaned

    def _polyline_signature(
        self, points: list[tuple[float, float]], precision_mm: float = 0.1
    ) -> tuple[tuple[int, int], ...]:
        quantized = tuple(
            (round(x / precision_mm), round(y / precision_mm)) for x, y in points
        )
        reversed_quantized = tuple(reversed(quantized))
        return min(quantized, reversed_quantized)

    def _float(self, value: str, label: str) -> float:
        try:
            return float(value.replace(",", "."))
        except ValueError as exc:
            raise ValueError(f"Invalid value for {label}: {value}") from exc

    def _transformed_polylines(
        self,
    ) -> tuple[
        list[list[tuple[float, float]]],
        tuple[float, float],
        tuple[float, float, float, float],
    ]:
        if not self.polylines_mm:
            raise ValueError("Load an SVG first.")

        media_width = self._float(self.width_var.get(), "width")
        media_height = self._float(self.height_var.get(), "height")
        box_offset = max(0.0, self._float(self.margin_var.get(), "peel box offset"))

        min_x, min_y, max_x, max_y = self.bounds_mm or (0.0, 0.0, 0.0, 0.0)
        width = max(max_x - min_x, 0.1)
        height = max(max_y - min_y, 0.1)

        normalized = [
            [(x - min_x, y - min_y) for x, y in line]
            for line in self.polylines_mm
        ]

        if self.rotate_var.get():
            normalized = [[(y, width - x) for x, y in line] for line in normalized]
            width, height = height, width

        if self.mirror_var.get():
            normalized = [[(width - x, y) for x, y in line] for line in normalized]

        available_width = max(media_width - box_offset * 2, 1.0)
        available_height = max(media_height - box_offset * 2, 1.0)

        scale = 1.0
        if self.fit_width_var.get():
            scale = min(available_width / width, available_height / height)

        origin_offset = box_offset if self.peel_box_var.get() else 0.0
        transformed = []
        for line in normalized:
            transformed.append(
                [(origin_offset + x * scale, origin_offset + y * scale) for x, y in line]
            )

        shape_width = width * scale
        shape_height = height * scale
        box_bounds = (
            0.0,
            0.0,
            shape_width + origin_offset * 2,
            shape_height + origin_offset * 2,
        )
        transformed = self._apply_output_flips(transformed, box_bounds)
        return transformed, (shape_width, shape_height), box_bounds

    def _peel_box_polyline(
        self, box_bounds: tuple[float, float, float, float]
    ) -> list[tuple[float, float]]:
        min_x, min_y, max_x, max_y = box_bounds
        return [
            (min_x, min_y),
            (max_x, min_y),
            (max_x, max_y),
            (min_x, max_y),
            (min_x, min_y),
        ]

    def _feed_after_target(
        self, box_bounds: tuple[float, float, float, float]
    ) -> tuple[float, float]:
        feed_after = max(0.0, self._float(self.feed_after_var.get(), "feed after cut"))
        min_x, min_y, max_x, max_y = box_bounds
        axis = self.feed_axis_var.get().strip().upper()
        direction = -1.0 if self.feed_direction_var.get().strip() == "-" else 1.0

        if axis == "Y":
            return min_x, max_y + (feed_after * direction)
        return max_x + (feed_after * direction), min_y

    def _apply_output_flips(
        self,
        polylines: list[list[tuple[float, float]]],
        box_bounds: tuple[float, float, float, float],
    ) -> list[list[tuple[float, float]]]:
        min_x, min_y, max_x, max_y = box_bounds
        transformed = polylines

        if self.flip_x_var.get():
            transformed = [[(max_x - (x - min_x), y) for x, y in line] for line in transformed]

        if self.flip_y_var.get():
            transformed = [[(x, max_y - (y - min_y)) for x, y in line] for line in transformed]

        return transformed

    def redraw_preview(self) -> None:
        self.canvas.delete("all")
        if not self.polylines_mm:
            self.canvas.create_text(
                self.canvas.winfo_width() / 2,
                self.canvas.winfo_height() / 2,
                text="Load / drag & drop an SVG to begin",
                fill="#c9d1d9",
                font=("Segoe UI", 15),
            )
            return

        try:
            polylines, (shape_width, shape_height), box_bounds = self._transformed_polylines()
        except Exception as exc:  # noqa: BLE001
            self.status_var.set(str(exc))
            return

        canvas_w = max(self.canvas.winfo_width(), 10)
        canvas_h = max(self.canvas.winfo_height(), 10)
        padding = 24
        min_x, min_y, max_x, max_y = box_bounds if self.peel_box_var.get() else (0.0, 0.0, shape_width, shape_height)
        preview_width = max(max_x - min_x, 1.0)
        preview_height = max(max_y - min_y, 1.0)
        scale = min((canvas_w - padding * 2) / preview_width, (canvas_h - padding * 2) / preview_height)

        if self.peel_box_var.get():
            self.canvas.create_rectangle(
                padding,
                padding,
                padding + (max_x - min_x) * scale,
                padding + (max_y - min_y) * scale,
                outline="#4f6478",
                dash=(5, 5),
            )

        for line in polylines:
            flat_points = []
            for x, y in line:
                flat_points.extend(
                    [padding + (x - min_x) * scale, padding + (y - min_y) * scale]
                )
            if len(flat_points) >= 4:
                self.canvas.create_line(
                    *flat_points,
                    fill="#7ee787",
                    width=1.6,
                    capstyle=tk.ROUND,
                    joinstyle=tk.ROUND,
                )

        self.status_var.set(
            f"Preview updated. Shape size: {shape_width:.1f} x {shape_height:.1f} mm."
        )

    def _build_hpgl(self) -> str:
        polylines, _, box_bounds = self._transformed_polylines()
        units_per_mm = self._float(self.units_var.get(), "HPGL units/mm")
        overcut_mm = max(0.0, self._float(self.overcut_var.get(), "overcut"))
        feed_after_mm = max(0.0, self._float(self.feed_after_var.get(), "feed after cut"))

        commands = ["IN;", "PA;", "SP1;"]
        for line in polylines:
            if len(line) < 2:
                continue
            start_x, start_y = line[0]
            commands.append(f"PU{round(start_x * units_per_mm)},{round(start_y * units_per_mm)};")

            draw_points = list(line[1:])
            if overcut_mm > 0:
                extended = self._extend_last_segment(line[-2], line[-1], overcut_mm)
                draw_points.append(extended)

            for x, y in draw_points:
                commands.append(f"PD{round(x * units_per_mm)},{round(y * units_per_mm)};")
            commands.append("PU;")

        if self.peel_box_var.get():
            peel_box = self._peel_box_polyline(box_bounds)
            start_x, start_y = peel_box[0]
            commands.append(f"PU{round(start_x * units_per_mm)},{round(start_y * units_per_mm)};")
            for x, y in peel_box[1:]:
                commands.append(f"PD{round(x * units_per_mm)},{round(y * units_per_mm)};")
            commands.append("PU;")

        if feed_after_mm > 0:
            feed_x, feed_y = self._feed_after_target(box_bounds)
            commands.append(f"PU{round(feed_x * units_per_mm)},{round(feed_y * units_per_mm)};")

        commands.append("SP0;")
        return "".join(commands)

    def _extend_last_segment(
        self, point_a: tuple[float, float], point_b: tuple[float, float], extra_mm: float
    ) -> tuple[float, float]:
        dx = point_b[0] - point_a[0]
        dy = point_b[1] - point_a[1]
        length = math.hypot(dx, dy)
        if length == 0:
            return point_b
        scale = extra_mm / length
        return point_b[0] + dx * scale, point_b[1] + dy * scale

    def save_hpgl(self) -> None:
        if not self.polylines_mm:
            messagebox.showinfo("No SVG", "Load an SVG first.")
            return

        try:
            hpgl = self._build_hpgl()
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("HPGL Error", str(exc))
            return

        target = filedialog.asksaveasfilename(
            title="Save HPGL",
            defaultextension=".plt",
            filetypes=[("Plotter files", "*.plt"), ("HPGL", "*.hpgl"), ("All files", "*.*")],
        )
        if not target:
            return

        Path(target).write_text(hpgl, encoding="ascii", errors="ignore")
        self.status_var.set(f"HPGL saved to {target}")

    def send_to_plotter(self) -> None:
        if not self.polylines_mm:
            messagebox.showinfo("No SVG", "Load an SVG first.")
            return

        try:
            hpgl = self._build_hpgl()
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Settings Error", str(exc))
            return

        transport = self._resolve_transport()
        if transport == "serial":
            if not self.port_var.get():
                messagebox.showinfo("No Port", "Choose a COM port first.")
                return
            try:
                baudrate = int(self._float(self.baud_var.get(), "baudrate"))
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("Baud Rate Error", str(exc))
                return

            device = self.port_var.get().split(" - ", 1)[0]
            self.status_var.set(f"Sending to {device} via serial...")
            thread = threading.Thread(
                target=self._send_serial_worker,
                args=(device, baudrate, hpgl),
                daemon=True,
            )
            thread.start()
            return

        if transport == "printer":
            if not self.printer_var.get():
                messagebox.showinfo(
                    "No Printer Queue",
                    "Choose a USB printer queue first or create one with 'Create USB Queue'.",
                )
                return

            printer_name = self.printer_var.get().split(" [", 1)[0]
            self.status_var.set(
                f"Sending to {printer_name} via Windows USB/Printer..."
            )
            thread = threading.Thread(
                target=self._send_printer_worker,
                args=(printer_name, hpgl),
                daemon=True,
            )
            thread.start()
            return

        messagebox.showerror(
            "No Transport Selected",
            "No usable serial or USB/printer connection is selected.",
        )

    def _resolve_transport(self) -> str | None:
        mode = self.transport_var.get()
        if mode == "Serial (COM)":
            return "serial"
        if mode == "Windows USB/Printer":
            return "printer"
        if self.printer_var.get():
            return "printer"
        if self.port_var.get():
            return "serial"
        return None

    def _clear_printer_jobs(self, printer_name: str) -> None:
        handle = None
        try:
            handle = win32print.OpenPrinter(printer_name)
            try:
                jobs = win32print.EnumJobs(handle, 0, 999, 1)
            except Exception:
                jobs = []

            for job in jobs:
                try:
                    win32print.SetJob(handle, job["JobId"], 0, None, win32print.JOB_CONTROL_DELETE)
                except Exception:
                    continue
        finally:
            if handle is not None:
                try:
                    win32print.ClosePrinter(handle)
                except Exception:
                    pass

    def _clear_openvinylcutter_queues(self) -> None:
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        try:
            printers = win32print.EnumPrinters(flags, None, 2)
        except Exception:
            return

        for info in printers:
            printer_name = info.get("pPrinterName") or ""
            if "OpenVinylCutter RAW" in printer_name:
                self._clear_printer_jobs(printer_name)

    def _send_serial_worker(self, device: str, baudrate: int, hpgl: str) -> None:
        try:
            with serial.Serial(
                port=device,
                baudrate=baudrate,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=2,
                write_timeout=10,
                xonxoff=False,
                rtscts=False,
                dsrdtr=False,
            ) as connection:
                payload = hpgl.encode("ascii", errors="ignore")
                connection.reset_output_buffer()
                connection.write(payload)
                connection.flush()
        except Exception as exc:  # noqa: BLE001
            self.root.after(
                0,
                lambda: messagebox.showerror("Send Failed", f"{device}\n\n{exc}"),
            )
            self.root.after(
                0, lambda: self.status_var.set(f"Sending to {device} failed.")
            )
            return

        self.root.after(
            0,
            lambda: self.status_var.set(
                f"Plot job sent to {device} at {baudrate} baud."
            ),
        )

    def _send_printer_worker(self, printer_name: str, hpgl: str) -> None:
        handle = None
        try:
            self._clear_openvinylcutter_queues()
            handle = win32print.OpenPrinter(printer_name)
            win32print.StartDocPrinter(handle, 1, ("OpenVinylCutter Plot Job", None, "RAW"))
            win32print.StartPagePrinter(handle)
            win32print.WritePrinter(handle, hpgl.encode("ascii", errors="ignore"))
            win32print.EndPagePrinter(handle)
            win32print.EndDocPrinter(handle)
        except Exception as exc:  # noqa: BLE001
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "USB/Printer Send Failed", f"{printer_name}\n\n{exc}"
                ),
            )
            self.root.after(
                0,
                lambda: self.status_var.set(
                    f"Sending to {printer_name} via USB/Printer failed."
                ),
            )
            return
        finally:
            if handle is not None:
                try:
                    win32print.ClosePrinter(handle)
                except Exception:
                    pass

        self.root.after(
            0,
            lambda: self.status_var.set(
                f"Plot job sent to {printer_name} via Windows USB/Printer."
            ),
        )


def main() -> None:
    root = tk.Tk()
    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = SvgPlotterApp(root)
    if len(sys.argv) > 1:
        startup_path = Path(sys.argv[1]).expanduser()
        if startup_path.exists() and startup_path.suffix.lower() == ".svg":
            root.after(50, lambda: app.open_svg_on_startup(startup_path))
    root.mainloop()


if __name__ == "__main__":
    main()
