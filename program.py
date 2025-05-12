import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
import openpyxl
from openpyxl import Workbook
from datetime import datetime
import json
import os

class EmployeeSchedulerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Распределение сотрудников")
        
        self.employees = []
        self.schedule = {}
        self.hall_names = {}
        self.days_of_week = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        
        self.load_data()
        self.create_widgets()
        self.create_schedule_tab()
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def load_data(self):
        if os.path.exists("scheduler_data.json"):
            try:
                with open("scheduler_data.json", "r") as f:
                    data = json.load(f)
                    self.employees = data.get("employees", [])
                    self.schedule = data.get("schedule", {})
                    self.hall_names = data.get("hall_names", {})
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка загрузки данных: {str(e)}")

    def save_data(self):
        try:
            data = {
                "employees": self.employees,
                "schedule": self.schedule,
                "hall_names": self.hall_names
            }
            with open("scheduler_data.json", "w") as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка сохранения данных: {str(e)}")

    def on_closing(self):
        self.save_data()
        self.root.destroy()

    def create_widgets(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)

        self.employee_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.employee_frame, text="Сотрудники")
        self.create_employee_tab()

        self.schedule_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.schedule_frame, text="Расписание залов")

    def create_employee_tab(self):
        input_frame = ttk.LabelFrame(self.employee_frame, text="Добавить сотрудника")
        input_frame.pack(padx=10, pady=10, fill="x")

        # Поля ввода для сотрудников
        ttk.Label(input_frame, text="Имя:").grid(row=0, column=0, padx=5, pady=2)
        self.name_entry = ttk.Entry(input_frame)
        self.name_entry.grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(input_frame, text="День:").grid(row=1, column=0, padx=5, pady=2)
        self.day_combobox = ttk.Combobox(input_frame, values=self.days_of_week)
        self.day_combobox.grid(row=1, column=1, padx=5, pady=2)
        
        ttk.Label(input_frame, text="Часы (через запятую):").grid(row=2, column=0, padx=5, pady=2)
        self.hours_entry = ttk.Entry(input_frame)
        self.hours_entry.grid(row=2, column=1, padx=5, pady=2)
        
        ttk.Label(input_frame, text="Макс. часов:").grid(row=3, column=0, padx=5, pady=2)
        self.max_hours_spinbox = ttk.Spinbox(input_frame, from_=1, to=24)
        self.max_hours_spinbox.grid(row=3, column=1, padx=5, pady=2)
        
        ttk.Button(input_frame, text="Добавить", command=self.add_employee).grid(row=4, columnspan=2, pady=5)
        
        # Таблица сотрудников
        self.tree = ttk.Treeview(self.employee_frame, columns=("Name", "Day", "Hours", "Max"), show="headings")
        self.tree.heading("Name", text="Имя")
        self.tree.heading("Day", text="День")
        self.tree.heading("Hours", text="Часы")
        self.tree.heading("Max", text="Макс. часов")
        self.tree.pack(padx=10, pady=10, fill="both", expand=True)
        
        # Кнопки управления
        btn_frame = ttk.Frame(self.employee_frame)
        btn_frame.pack(fill="x", padx=10, pady=5)
        ttk.Button(btn_frame, text="Удалить", command=self.delete_employee).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="Сохранить в Excel", command=self.save_schedule).pack(side=tk.RIGHT)

    def create_schedule_tab(self):
        # Фрейм для управления расписанием
        control_frame = ttk.LabelFrame(self.schedule_frame, text="Настройка расписания")
        control_frame.pack(padx=10, pady=10, fill="x")

        # Элементы управления
        ttk.Label(control_frame, text="Дата:").grid(row=0, column=0, padx=5, pady=2)
        self.date_entry = DateEntry(control_frame, date_pattern='dd.mm.yyyy')
        self.date_entry.grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(control_frame, text="Час:").grid(row=0, column=2, padx=5, pady=2)
        self.hour_spinbox = ttk.Spinbox(control_frame, from_=0, to=23)
        self.hour_spinbox.grid(row=0, column=3, padx=5, pady=2)
        
        ttk.Label(control_frame, text="Кол-во залов:").grid(row=0, column=4, padx=5, pady=2)
        self.halls_spinbox = ttk.Spinbox(control_frame, from_=1, to=10)
        self.halls_spinbox.grid(row=0, column=5, padx=5, pady=2)
        
        ttk.Label(control_frame, text="Названия залов:").grid(row=1, column=0, padx=5, pady=2)
        self.hall_names_entry = ttk.Entry(control_frame)
        self.hall_names_entry.grid(row=1, column=1, columnspan=5, sticky="we", padx=5, pady=2)
        
        # Кнопки
        btn_frame = ttk.Frame(control_frame)
        btn_frame.grid(row=2, column=0, columnspan=6, pady=5)
        ttk.Button(btn_frame, text="Добавить час", command=self.add_hour).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Удалить", command=self.remove_hour).pack(side=tk.LEFT, padx=5)

        # Таблица расписания
        self.schedule_tree = ttk.Treeview(self.schedule_frame, 
                                        columns=("Date", "Hour", "Halls", "Names"), 
                                        show="headings")
        self.schedule_tree.heading("Date", text="Дата")
        self.schedule_tree.heading("Hour", text="Час")
        self.schedule_tree.heading("Halls", text="Кол-во залов")
        self.schedule_tree.heading("Names", text="Названия")
        self.schedule_tree.pack(padx=10, pady=10, fill="both", expand=True)
        
        self.update_schedule_treeview()

    def add_employee(self):
        try:
            name = self.name_entry.get()
            day = self.day_combobox.get()
            hours = list(map(int, self.hours_entry.get().split(',')))
            max_hours = int(self.max_hours_spinbox.get())
            
            if not all([name, day, hours, max_hours]):
                messagebox.showwarning("Ошибка", "Заполните все поля")
                return
                
            self.employees.append({
                'name': name,
                'availability': {day: {'hours_available': hours, 'max_hours': max_hours}},
                'assigned_hours': []
            })
            self.update_treeview()
            self.clear_inputs()
            
        except ValueError:
            messagebox.showerror("Ошибка", "Неверный формат данных")

    def delete_employee(self):
        selected = self.tree.selection()
        if selected:
            self.employees.pop(self.tree.index(selected[0]))
            self.update_treeview()

    def clear_inputs(self):
        self.name_entry.delete(0, tk.END)
        self.hours_entry.delete(0, tk.END)
        self.max_hours_spinbox.delete(0, tk.END)
        self.day_combobox.set('')

    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        for emp in self.employees:
            for day in emp['availability']:
                hours = ','.join(map(str, emp['availability'][day]['hours_available']))
                max_h = emp['availability'][day]['max_hours']
                self.tree.insert("", "end", values=(emp['name'], day, hours, max_h))

    def add_hour(self):
        try:
            date = self.date_entry.get_date().strftime("%d.%m.%Y")
            day_of_week = self.date_entry.get_date().strftime("%A")
            hour = int(self.hour_spinbox.get())
            halls = int(self.halls_spinbox.get())
            names = [n.strip() for n in self.hall_names_entry.get().split(',')[:halls]]
            
            key = f"{date} ({day_of_week})"
            
            if key not in self.schedule:
                self.schedule[key] = []
                self.hall_names[key] = {}
            
            self.schedule[key].append({'hour': hour, 'halls': halls})
            self.hall_names[key][hour] = names + [f"Зал {i+1}" for i in range(len(names), halls)]
            
            self.update_schedule_treeview()
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка ввода: {str(e)}")

    def remove_hour(self):
        selected = self.schedule_tree.selection()
        if selected:
            item = self.schedule_tree.item(selected[0])
            key, hour = item['values'][0], int(item['values'][1])
            self.schedule[key] = [h for h in self.schedule[key] if h['hour'] != hour]
            if not self.schedule[key]: 
                del self.schedule[key]
                del self.hall_names[key]
            else:
                del self.hall_names[key][hour]
            self.update_schedule_treeview()

    def update_schedule_treeview(self):
        self.schedule_tree.delete(*self.schedule_tree.get_children())
        for key in self.schedule:
            for hour_info in self.schedule[key]:
                hour = hour_info['hour']
                halls = hour_info['halls']
                names = ', '.join(self.hall_names[key].get(hour, []))
                self.schedule_tree.insert("", "end", values=(key, hour, halls, names))

    def distribute_employees(self):
        for emp in self.employees:
            emp['assigned_hours'] = []
        
        for key in self.schedule:
            day_name = key.split('(')[-1].rstrip(')')
            for hour_info in sorted(self.schedule[key], key=lambda x: (-x['halls'], x['hour'])):
                hour = hour_info['hour']
                required = hour_info['halls']
                
                candidates = []
                for emp in self.employees:
                    if day_name in emp['availability']:
                        avail = emp['availability'][day_name]
                        if hour in avail['hours_available']:
                            assigned = len([h for h in emp['assigned_hours'] if h['day'] == key])
                            if assigned < avail['max_hours']:
                                candidates.append(emp)
                
                candidates.sort(key=lambda e: (
                    e['availability'][day_name]['max_hours'] - 
                    len([h for h in e['assigned_hours'] if h['day'] == key])
                ), reverse=True)
                
                for emp in candidates[:required]:
                    emp['assigned_hours'].append({'day': key, 'hour': hour})

    def save_schedule(self):
        self.distribute_employees()
        
        # Собираем данные по сотрудникам
        employee_assignments = {}
        for key in self.schedule:
            date_str, day_name = key.split(' (')
            day_name = day_name.rstrip(')')
            
            for hour_info in self.schedule[key]:
                hour = hour_info['hour']
                halls = hour_info['halls']
                hall_names = self.hall_names[key].get(hour, [])
                
                # Получаем список сотрудников для этого часа
                assigned_employees = []
                for emp in self.employees:
                    if any(assgn['day'] == key and assgn['hour'] == hour 
                        for assgn in emp['assigned_hours']):
                        assigned_employees.append(emp['name'])
                
                # Распределяем сотрудников по залам
                for i in range(halls):
                    hall_name = hall_names[i] if i < len(hall_names) else f"Зал {i+1}"
                    emp_name = assigned_employees[i] if i < len(assigned_employees) else "Не назначен"
                    
                    if emp_name not in employee_assignments:
                        employee_assignments[emp_name] = []
                    
                    employee_assignments[emp_name].append({
                        'date': date_str,
                        'day': day_name,
                        'hour': hour,
                        'hall': hall_name
                    })

        # Создаем Excel-файл
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel файлы", "*.xlsx")],
            title="Сохранить расписание"
        )
        
        if not file_path: 
            return
        
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "График работы"
            
            # Заголовки
            headers = ["Сотрудник", "Дата", "День", "Зал", "Часы работы"]
            ws.append(headers)
            
            # Обрабатываем данные для каждого сотрудника
            for emp_name, assignments in employee_assignments.items():
                # Группируем по дате и залу
                grouped = {}
                for assignment in assignments:
                    key = (assignment['date'], assignment['hall'])
                    if key not in grouped:
                        grouped[key] = []
                    grouped[key].append(assignment['hour'])
                
                # Формируем временные интервалы
                for (date, hall), hours in grouped.items():
                    hours = sorted(hours)
                    time_ranges = []
                    start = end = hours[0]
                    
                    for h in hours[1:]:
                        if h == end + 1:
                            end = h
                        else:
                            time_ranges.append(f"{start:02d}:00-{end+1:02d}:00")
                            start = end = h
                    time_ranges.append(f"{start:02d}:00-{end+1:02d}:00")
                    
                    # Записываем в таблицу
                    for time_range in time_ranges:
                        ws.append([
                            emp_name,
                            date,
                            hall,
                            time_range
                        ])
            
            # Настраиваем ширину столбцов
            for col in ws.columns:
                max_length = max(len(str(cell.value)) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = max_length + 2
            
            wb.save(file_path)
            messagebox.showinfo("Успех", f"Файл сохранен:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка сохранения: {str(e)}")
if __name__ == "__main__":
    root = tk.Tk()
    app = EmployeeSchedulerApp(root)
    root.mainloop()
