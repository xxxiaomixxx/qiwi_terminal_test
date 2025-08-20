# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import serial
import time
import threading
from queue import Queue
from PIL import Image, ImageTk, ImageOps

# -----------------------------------------------------------------------------
# КЛАСС ДЛЯ УПРАВЛЕНИЯ ПРИНТЕРОМ
# -----------------------------------------------------------------------------
class CitizenPPU700:
    def __init__(self, port, baudrate=19200, timeout=1):
        self.device = None
        self.error_message = None
        try:
            self.device = serial.Serial(
                port=port, baudrate=baudrate, bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                timeout=timeout, dsrdtr=True
            )
        except serial.SerialException as e:
            self.error_message = f"Не удалось открыть порт принтера {port}.\n\n{e}"

    def is_connected(self):
        return self.device is not None and self.device.is_open

    def _write(self, data):
        if self.is_connected(): self.device.write(data); self.device.flush()

    def initialize(self): self._write(b'\x1b\x40')
    
    def set_code_page(self, codepage='cp437'):
        """Устанавливает кодовую страницу на принтере."""
        codepage_map = {
            'cp437': 0,  # Стандартная страница для команд
            'cp866': 17  # Русская страница для текста
        }
        page_number = codepage_map.get(codepage, 0)
        self._write(b'\x1b\x74' + bytes([page_number]))

    def text(self, txt, encoding='cp866'): self._write(txt.encode(encoding))
    def feed(self, lines=1): self._write(b'\n' * lines)
    def set_alignment(self, align='left'): self._write(b'\x1b\x61' + bytes([{'left': 0, 'center': 1, 'right': 2}.get(align, 0)]))
    def cut(self): self._write(b'\x1d\x56\x00')
    def close(self):
        if self.is_connected(): self.device.close()
        
    def print_barcode(self, data, barcode_type='CODE128'):
        barcode_map = {'CODE39': 69, 'CODE128': 73, 'EAN13': 67}
        m = barcode_map.get(barcode_type)
        if m is None: return

        self._write(b'\x1d\x68\x50') # Установить высоту штрихкода
        self._write(b'\x1d\x48\x02') # Печатать текст под штрихкодом
        encoded_data = data.encode('ascii')
        n = len(encoded_data)
        self._write(b'\x1d\x6b' + bytes([m, n]) + encoded_data)

    def print_image(self, file_path, max_width=576):
        try:
            img = Image.open(file_path)
            if img.width > max_width:
                ratio = max_width / img.width
                new_height = int(img.height * ratio)
                img = img.resize((max_width, new_height), Image.Resampling.LANCZOS)
            img = img.convert('L')
            img = ImageOps.invert(img)
            img = img.convert('1', dither=Image.Dither.FLOYDSTEINBERG)
            width, height = img.size
            width_bytes = (width + 7) // 8
            cmd_header = b'\x1d\x76\x30\x00' + bytes([width_bytes & 0xFF, (width_bytes >> 8) & 0xFF, height & 0xFF, (height >> 8) & 0xFF])
            self.set_alignment('center')
            self._write(cmd_header + img.tobytes())
            self.feed(1)
            return True
        except Exception as e:
            self.error_message = f"Не удалось обработать файл изображения.\n\n{e}"
            return False

# -----------------------------------------------------------------------------
# КЛАСС ДЛЯ УПРАВЛЕНИЯ КУПЮРОПРИЕМНИКОМ
# -----------------------------------------------------------------------------
class CashCodeAcceptor:
    CMD_RESET = bytes([0x02, 0x03, 0x06, 0x30, 0x41, 0xB3])
    CMD_ACK = bytes([0x02, 0x03, 0x06, 0x00, 0xC2, 0x82])
    CMD_POLL = bytes([0x02, 0x03, 0x06, 0x33, 0xDA, 0x81])
    CMD_ENABLE_ALL = bytes([0x02, 0x03, 0x0C, 0x34, 0x00, 0x30, 0xFC, 0x00, 0x00, 0x00, 0xD9, 0x38])
    DENOMINATIONS = {0x07: 5000, 0x0D: 2000, 0x06: 1000, 0x05: 500, 0x0C: 200, 0x04: 100, 0x03: 50, 0x02: 10}

    def __init__(self, port, gui_queue):
        self.port = port
        self.gui_queue = gui_queue
        self.device = None
        self.is_running = False
        self.thread = None

    def start(self):
        self.is_running = True
        self.thread = threading.Thread(target=self._run_loop, daemon=True)
        self.thread.start()

    def stop(self):
        self.is_running = False
        if self.thread:
            self.thread.join(timeout=1)

    def reset_device(self):
        """Отправляет команду сброса и останавливает цикл работы."""
        if self.device and self.device.is_open:
            try:
                self.gui_queue.put(("status", "Перезагрузка устройства..."))
                self.device.write(self.CMD_RESET)
                time.sleep(0.1)
                self.is_running = False # Сигнал для остановки цикла и закрытия порта
                self.gui_queue.put(("status", "Перезагружено. Нажмите 'Инициализация'."))
            except serial.SerialException:
                self.gui_queue.put(("status", "Ошибка при перезагрузке"))
        else:
            self.gui_queue.put(("status", "Устройство не подключено"))

    def _run_loop(self):
        try:
            self.device = serial.Serial(self.port, 9600, timeout=0.1)
            self.gui_queue.put(("status", "Инициализация..."))
            self.device.write(self.CMD_RESET)
            time.sleep(2)
            self.device.flushInput()
            self.device.write(self.CMD_ENABLE_ALL)
            time.sleep(0.2)
            self.device.flushInput()
            self.gui_queue.put(("status", "Готов к приему купюр"))
        except serial.SerialException:
            self.gui_queue.put(("status", f"Ошибка порта: {self.port}"))
            return

        while self.is_running:
            try:
                self.device.write(self.CMD_POLL)
                time.sleep(0.1)
                if self.device.inWaiting() > 0:
                    response = self.device.read(6)
                    if len(response) >= 6 and response[3] == 0x81:
                        denomination = self.DENOMINATIONS.get(response[4])
                        if denomination:
                            self.gui_queue.put(("bill", denomination))
                            self.device.write(self.CMD_ACK)
                            time.sleep(0.1)
                            self.device.write(self.CMD_ENABLE_ALL)
            except serial.SerialException:
                self.gui_queue.put(("status", "Ошибка связи"))
                break
        
        if self.device and self.device.is_open:
            self.device.close()
            print(f"Порт купюроприемника {self.port} закрыт.")

# -----------------------------------------------------------------------------
# КЛАСС ГРАФИЧЕСКОГО ИНТЕРФЕЙСА (GUI)
# -----------------------------------------------------------------------------
class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Панель управления Qiwi терминалом")
        self.minsize(800, 600)

        self.printer_thread = None
        self.acceptor_instance = None
        self.gui_queue = Queue()
        self.total_sum = 0
        
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))

        self.create_printer_widgets(left_frame)
        self.create_acceptor_widgets(right_frame)

        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.process_queue()

    def create_printer_widgets(self, parent):
        printer_frame = ttk.LabelFrame(parent, text="Принтер", padding=10)
        printer_frame.pack(fill="both", expand=True)
        
        port_frame = ttk.LabelFrame(printer_frame, text="Подключение", padding=10)
        port_frame.pack(padx=10, pady=10, fill="x")
        ttk.Label(port_frame, text="COM Порт:").pack(side="left", padx=5)
        self.printer_port_entry = ttk.Entry(port_frame); self.printer_port_entry.insert(0, "COM3")
        self.printer_port_entry.pack(side="left", fill="x", expand=True)

        text_frame = ttk.LabelFrame(printer_frame, text="Печать текста", padding=10)
        text_frame.pack(padx=10, pady=5, fill="both", expand=True)
        self.text_input = tk.Text(text_frame, height=8, wrap="word"); self.text_input.insert("1.0", "Привет, мир!")
        self.text_input.pack(fill="both", expand=True, pady=5)
        ttk.Button(text_frame, text="Печать текста", command=self.print_text).pack(fill="x")

        code_frame = ttk.LabelFrame(printer_frame, text="Печать штрихкода / QR-кода (BETA)", padding=10)
        code_frame.pack(padx=10, pady=5, fill="x")
        self.code_input = ttk.Entry(code_frame); self.code_input.insert(0, "1234567890")
        self.code_input.pack(fill="x", pady=5)
        btn_frame = ttk.Frame(code_frame); btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="Печать штрихкода (Code128)", command=self.print_barcode).pack(side="left", expand=True, fill="x", padx=(0,5))
        ttk.Button(btn_frame, text="Печать QR-кода", command=lambda: messagebox.showinfo("Инфо", "Функция QR-кода не реализована.")).pack(side="left", expand=True, fill="x")

        img_frame = ttk.LabelFrame(printer_frame, text="Печать изображения", padding=10)
        img_frame.pack(padx=10, pady=5, fill="x")
        ttk.Button(img_frame, text="Выбрать и напечатать изображение...", command=self.select_and_print_image).pack(fill="x")

        control_frame = ttk.Frame(printer_frame, padding=10)
        control_frame.pack(padx=10, pady=10, fill="x")
        ttk.Button(control_frame, text="Отрезать чек", command=self.cut_paper).pack(fill="x")

    def create_acceptor_widgets(self, parent):
        acceptor_frame = ttk.LabelFrame(parent, text="Купюроприемник", padding=10)
        acceptor_frame.pack(fill="both", expand=True)

        port_frame = ttk.LabelFrame(acceptor_frame, text="Подключение", padding=10)
        port_frame.pack(padx=10, pady=10, fill="x")
        ttk.Label(port_frame, text="COM Порт:").pack(side="left", padx=5)
        self.acceptor_port_entry = ttk.Entry(port_frame); self.acceptor_port_entry.insert(0, "COM4")
        self.acceptor_port_entry.pack(side="left", fill="x", expand=True)

        status_frame = ttk.LabelFrame(acceptor_frame, text="Статус", padding=10)
        status_frame.pack(padx=10, pady=10, fill="both", expand=True)
        
        self.status_var = tk.StringVar(value="Ожидание...")
        self.last_bill_var = tk.StringVar(value="Последняя распознанная купюра: -")
        self.total_sum_var = tk.StringVar(value="Всего распознано: 0 руб")

        ttk.Label(status_frame, textvariable=self.status_var, font=('Helvetica', 10, 'italic')).pack(anchor="w", pady=5)
        ttk.Label(status_frame, textvariable=self.last_bill_var, font=('Helvetica', 12)).pack(anchor="w", pady=5)
        ttk.Label(status_frame, textvariable=self.total_sum_var, font=('Helvetica', 12, 'bold')).pack(anchor="w", pady=5)

        acceptor_btn_frame = ttk.Frame(acceptor_frame)
        acceptor_btn_frame.pack(fill="x", padx=10, pady=10)
        
        self.init_button = ttk.Button(acceptor_btn_frame, text="Инициализация", command=self.start_acceptor)
        self.init_button.pack(fill="x", pady=(0, 5))
        
        self.reset_button = ttk.Button(acceptor_btn_frame, text="Перезагрузить", command=self.reset_acceptor)
        self.reset_button.pack(fill="x")

    def process_queue(self):
        try:
            message = self.gui_queue.get_nowait()
            msg_type, value = message
            if msg_type == "status":
                self.status_var.set(value)
            elif msg_type == "bill":
                self.total_sum += value
                self.last_bill_var.set(f"Последняя распознанная купюра: {value} руб")
                self.total_sum_var.set(f"Всего распознано: {self.total_sum} руб")
        except Exception:
            pass
        self.after(100, self.process_queue)

    def start_acceptor(self):
        if self.acceptor_instance:
            self.acceptor_instance.stop()
        port = self.acceptor_port_entry.get()
        self.total_sum = 0 
        self.last_bill_var.set("Последняя распознанная купюра: -")
        self.total_sum_var.set("Всего распознано: 0 руб")
        self.acceptor_instance = CashCodeAcceptor(port, self.gui_queue)
        self.acceptor_instance.start()

    def reset_acceptor(self):
        if self.acceptor_instance and self.acceptor_instance.is_running:
            self.acceptor_instance.reset_device()
        else:
            messagebox.showwarning("Внимание", "Купюроприемник не запущен. Сначала выполните инициализацию.")

    def run_printer_job(self, job_func, *args):
        printer = CitizenPPU700(self.printer_port_entry.get())
        if not printer.is_connected():
            self.after(0, messagebox.showerror, "Ошибка принтера", printer.error_message)
            return
        try:
            job_func(printer, *args)
            printer.feed(2)
        finally:
            printer.close()

    def start_printer_thread(self, job_func, *args):
        if self.printer_thread and self.printer_thread.is_alive():
            messagebox.showwarning("Принтер занят", "Дождитесь завершения предыдущей операции.")
            return
        self.printer_thread = threading.Thread(target=self.run_printer_job, args=(job_func, *args), daemon=True)
        self.printer_thread.start()

    def print_text(self):
        text_to_print = self.text_input.get("1.0", tk.END)
        def job(printer, text):
            printer.initialize()
            printer.set_code_page('cp866')
            printer.text(text)
        self.start_printer_thread(job, text_to_print)

    def print_barcode(self):
        data = self.code_input.get()
        def job(printer, data_to_print):
            printer.initialize()
            printer.set_code_page('cp437') 
            printer.print_barcode(data_to_print)
        if data: self.start_printer_thread(job, data)

    def select_and_print_image(self):
        file_path = filedialog.askopenfilename(filetypes=[("Изображения", "*.png *.jpg *.bmp")])
        def job(printer, path):
            printer.initialize()
            printer.set_code_page('cp437')
            printer.print_image(path)
        if file_path: self.start_printer_thread(job, file_path)

    def cut_paper(self):
        def job(printer):
            printer.initialize()
            printer.set_code_page('cp437')
            printer.cut()
        self.start_printer_thread(job)

    def on_closing(self):
        if self.acceptor_instance:
            self.acceptor_instance.stop()
        self.destroy()

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
