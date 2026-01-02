import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import re
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

class DarkDistrictAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Анализ отчетов ФЦМПО по районам")
        self.root.geometry("900x700")
        self.root.configure(bg='#2b2b2b')
        
        # Полный список районов для поиска
        self.districts = [
            "Аргун", "Ачхой-Мартановский", "Веденский", "Грозненский", "Грозный",
            "Гудермесский", "Гудермес", "Итум-Калинский", "Курчалоевский", "Надтеречный",
            "Наурский", "Ножай-Юртовский", "Серноводский", "Урус-Мартановский",
            "Шалинский", "Шаройский", "Шатойский", "Шелковской"
        ]
        
        self.setup_styles()
        self.setup_ui()
        self.results = None
        
    def setup_styles(self):
        """Настройка темных стилей"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Темная цветовая схема
        dark_bg = '#2b2b2b'
        dark_fg = '#ffffff'
        darker_bg = '#1e1e1e'
        accent_color = '#3a7ca5'
        
        style.configure('TFrame', background=dark_bg)
        style.configure('TLabel', background=dark_bg, foreground=dark_fg, font=('Arial', 10))
        style.configure('TButton', background=accent_color, foreground=dark_fg, 
                       font=('Arial', 10, 'bold'), borderwidth=0)
        style.map('TButton', background=[('active', '#2a6480')])
        
        style.configure('Title.TLabel', font=('Arial', 14, 'bold'), background=dark_bg, foreground=dark_fg)
        style.configure('Card.TFrame', background=darker_bg, relief='flat')
        
        # Стиль для Treeview (таблицы)
        style.configure("Dark.Treeview", 
                       background=darker_bg,
                       foreground=dark_fg,
                       fieldbackground=darker_bg,
                       borderwidth=0)
        style.configure("Dark.Treeview.Heading", 
                       background='#3d3d3d',
                       foreground=dark_fg,
                       relief='flat',
                       font=('Arial', 10, 'bold'))
        
    def setup_ui(self):
        # Главный контейнер
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок
        title_label = ttk.Label(main_frame, text="Анализ отчетов ФЦМПО по районам", 
                               style='Title.TLabel')
        title_label.pack(pady=(0, 20))
        
        # Фрейм загрузки файла
        file_frame = ttk.Frame(main_frame, style='Card.TFrame', padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(file_frame, text="Загрузка файла данных:").pack(anchor=tk.W)
        
        file_subframe = ttk.Frame(file_frame)
        file_subframe.pack(fill=tk.X, pady=(5, 0))
        
        self.load_btn = ttk.Button(file_subframe, text="Выбрать файл Excel", command=self.load_file)
        self.load_btn.pack(side=tk.LEFT)
        
        self.file_path_var = tk.StringVar(value="Файл не выбран")
        file_path_label = ttk.Label(file_subframe, textvariable=self.file_path_var, 
                                   background='#3d3d3d', foreground='#cccccc', 
                                   relief='solid', borderwidth=1, padding=(5, 2))
        file_path_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))
        
        # Фрейм настроек анализа
        settings_frame = ttk.Frame(main_frame, style='Card.TFrame', padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(settings_frame, text="Тип анализа:").pack(anchor=tk.W)
        
        self.analysis_var = tk.StringVar(value="score5")
        
        analysis_frame = ttk.Frame(settings_frame)
        analysis_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Radiobutton(analysis_frame, text="Балл нормированный к 5", 
                       variable=self.analysis_var, value="score5").pack(side=tk.LEFT, padx=(0, 20))
        ttk.Radiobutton(analysis_frame, text="Соответствие типовому меню", 
                       variable=self.analysis_var, value="menu_compliance").pack(side=tk.LEFT)
        
        # Фрейм управления
        control_frame = ttk.Frame(main_frame, style='Card.TFrame', padding="10")
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.analyze_btn = ttk.Button(control_frame, text="Анализировать", 
                                     command=self.analyze_file, state="disabled")
        self.analyze_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.save_btn = ttk.Button(control_frame, text="Сохранить результаты", 
                                  command=self.save_results, state="disabled")
        self.save_btn.pack(side=tk.LEFT)
        
        # Фрейм результатов
        results_frame = ttk.Frame(main_frame, style='Card.TFrame', padding="10")
        results_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(results_frame, text="Результаты анализа:").pack(anchor=tk.W, pady=(0, 10))
        
        # Таблица результатов
        tree_frame = ttk.Frame(results_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(tree_frame, columns=("district", "score", "count"), 
                                show="headings", style="Dark.Treeview", height=15)
        
        self.tree.heading("district", text="Район")
        self.tree.heading("score", text="Показатель")
        self.tree.heading("count", text="Количество записей")
        
        self.tree.column("district", width=250)
        self.tree.column("score", width=150)
        self.tree.column("count", width=120)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Статус бар
        self.status_var = tk.StringVar(value="Готов к работе")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief='sunken', background='#3d3d3d', 
                              foreground='#cccccc', padding=(5, 2))
        status_bar.pack(fill=tk.X, pady=(10, 0))
    
    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            self.file_path_var.set(file_path)
            self.analyze_btn.config(state="normal")
            self.status_var.set(f"Файл загружен: {os.path.basename(file_path)}")
    
    def analyze_file(self):
        file_path = self.file_path_var.get()
        if file_path == "Файл не выбран":
            messagebox.showerror("Ошибка", "Сначала выберите файл")
            return
        
        analysis_type = self.analysis_var.get()
        self.status_var.set("Анализ файла...")
        self.analyze_btn.config(state="disabled")
        self.root.update()
        
        try:
            df = pd.read_excel(file_path, header=None)
            
            if analysis_type == "score5":
                results = self.analyze_score5(df)
            else:
                results = self.analyze_menu_compliance(df)
            
            self.results = results
            self.display_results(results)
            self.save_btn.config(state="normal")
            self.status_var.set(f"Анализ завершен. Обработано {len(results)} районов")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при анализе файла: {str(e)}")
            self.status_var.set("Ошибка при анализе")
        
        finally:
            self.analyze_btn.config(state="normal")
    
    def analyze_score5(self, df):
        """Анализ балла нормированного к 5 (оригинальный алгоритм)"""
        district_data = {}
        current_district = None
        
        for i, row in df.iterrows():
            cell_value = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
            
            # Проверяем, является ли ячейка названием района
            for district in self.districts:
                if cell_value.strip() == district:
                    current_district = district
                    if district not in district_data:
                        district_data[district] = {
                            'score5_data': [],
                            'record_count': 0
                        }
                    break
            
            # Если мы находимся в блоке района и это строка с данными
            if current_district and i > 0:
                # Проверяем, является ли строка данными (дата в первом столбце)
                if re.match(r'\d{1,2}\.\d{1,2}\.\d{4}', str(cell_value)):
                    district_data[current_district]['record_count'] += 1
                    
                    # Поиск столбца с баллом нормированным к 5
                    score5_col = None
                    
                    # Поиск по заголовкам в предыдущей строке
                    for j in range(len(row)):
                        if i > 0 and pd.notna(df.iloc[i-1, j]):
                            header_val = str(df.iloc[i-1, j])
                            if "нормированный к5" in header_val.lower():
                                score5_col = j
                                break
                    
                    # Если не нашли по заголовкам, используем поиск по типу данных
                    if score5_col is None:
                        # Ищем числовой столбец с баллами от 0 до 5
                        for j in range(len(row)-1, max(len(row)-10, 0), -1):
                            try:
                                val = float(row.iloc[j])
                                if 0 <= val <= 5:
                                    score5_col = j
                                    break
                            except:
                                continue
                    
                    # Собираем данные балла нормированного к 5
                    if score5_col is not None and pd.notna(row.iloc[score5_col]):
                        try:
                            score_val = float(row.iloc[score5_col])
                            district_data[current_district]['score5_data'].append(score_val)
                        except:
                            pass
        
        # Расчет статистики
        results = []
        for district in self.districts:
            if district in district_data:
                data = district_data[district]
                if data['record_count'] > 0:
                    # Средний балл нормированный к 5 (в процентах)
                    score5_percent = 0
                    if data['score5_data']:
                        score5_percent = (sum(data['score5_data']) / len(data['score5_data'])) / 5 * 100
                    
                    results.append({
                        'district': district,
                        'score': round(score5_percent, 2),
                        'record_count': data['record_count'],
                        'analysis_type': 'score5'
                    })
            else:
                # Добавляем район с нулевыми результатами, если он не найден
                results.append({
                    'district': district,
                    'score': 0.0,
                    'record_count': 0,
                    'analysis_type': 'score5'
                })
        
        # Сортируем по проценту (от лучшего к худшему)
        results.sort(key=lambda x: x['score'], reverse=True)
        return results
    
    def analyze_menu_compliance(self, df):
        """Анализ соответствия типовому меню с расчетом общего процента по регионам"""
        district_data = {}
        all_compliance_data = []  # Все проценты по региону
        
        # Проходим по всем строкам для поиска данных по районам
        for i, row in df.iterrows():
            # Пропускаем пустые строки
            if pd.isna(row.iloc[0]):
                continue
                
            # Ищем строки с названиями учреждений (содержат "МБОУ", "СОШ" и т.д.)
            cell_value = str(row.iloc[0])
            if any(keyword in cell_value for keyword in ['МБОУ', 'СОШ', 'ГБОУ', 'Школа']):
                # Проверяем, есть ли в строке название района
                district_found = None
                for district in self.districts:
                    if district in (str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""):
                        district_found = district
                        break
                
                if district_found:
                    if district_found not in district_data:
                        district_data[district_found] = {
                            'compliance_data': [],
                            'record_count': 0
                        }
                    
                    # Ищем столбцы с соответствием типовому меню (последние 2 столбца)
                    for j in range(len(row)-2, len(row)):
                        if j >= 0 and pd.notna(row.iloc[j]):
                            try:
                                # Ищем процент соответствия (значение 0-100)
                                compliance_val = float(row.iloc[j])
                                if 0 <= compliance_val <= 100:
                                    district_data[district_found]['compliance_data'].append(compliance_val)
                                    district_data[district_found]['record_count'] += 1
                                    all_compliance_data.append(compliance_val)
                                    break
                            except:
                                continue
        
        # Расчет общих показателей по региону
        region_stats = self.calculate_region_stats(all_compliance_data)
        
        # Расчет статистики по районам
        results = []
        for district in self.districts:
            if district in district_data:
                data = district_data[district]
                if data['record_count'] > 0:
                    # Средний процент соответствия
                    avg_compliance = 0
                    if data['compliance_data']:
                        avg_compliance = sum(data['compliance_data']) / len(data['compliance_data'])
                    
                    results.append({
                        'district': district,
                        'score': round(avg_compliance, 2),
                        'record_count': data['record_count'],
                        'analysis_type': 'menu_compliance'
                    })
            else:
                # Добавляем район с нулевыми результатами, если он не найден
                results.append({
                    'district': district,
                    'score': 0.0,
                    'record_count': 0,
                    'analysis_type': 'menu_compliance'
                })
        
        # Добавляем строки с общими показателями по региону
        results.append({
            'district': 'ОБЩИЙ ПОКАЗАТЕЛЬ ПО РЕГИОНУ (100%)',
            'score': round(region_stats['perfect_percentage'], 2),
            'record_count': len(all_compliance_data),
            'analysis_type': 'menu_compliance',
            'is_region_total': True,
            'category': '100%'
        })
        
        results.append({
            'district': 'ОБЩИЙ ПОКАЗАТЕЛЬ ПО РЕГИОНУ (75-100%)',
            'score': round(region_stats['good_percentage'], 2),
            'record_count': len(all_compliance_data),
            'analysis_type': 'menu_compliance',
            'is_region_total': True,
            'category': '75-100%'
        })
        
        # Сортируем по проценту соответствия (от лучшего к худшему), но общие показатели оставляем в конце
        regular_results = [r for r in results if not r.get('is_region_total', False)]
        regular_results.sort(key=lambda x: x['score'], reverse=True)
        
        # Добавляем общие показатели в конец
        region_results = [r for r in results if r.get('is_region_total', False)]
        return regular_results + region_results
    
    def calculate_region_stats(self, compliance_data):
        """Расчет статистики по региону"""
        if not compliance_data:
            return {'perfect_percentage': 0, 'good_percentage': 0}
        
        # Процент организаций с 100% соответствием
        perfect_compliance_count = sum(1 for x in compliance_data if x == 100)
        perfect_percentage = (perfect_compliance_count / len(compliance_data)) * 100
        
        # Процент организаций с 75-100% соответствием
        good_compliance_count = sum(1 for x in compliance_data if 75 <= x <= 100)
        good_percentage = (good_compliance_count / len(compliance_data)) * 100
        
        return {
            'perfect_percentage': perfect_percentage,
            'good_percentage': good_percentage
        }
    
    def display_results(self, results):
        # Очищаем таблицу
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Обновляем заголовки в зависимости от типа анализа
        analysis_type = self.analysis_var.get()
        if analysis_type == "score5":
            self.tree.heading("score", text="Балл нормированный к 5 (%)")
        else:
            self.tree.heading("score", text="Соответствие меню (%)")
        
        # Заполняем таблицу результатами
        for result in results:
            score_text = f"{result['score']}%"
            
            # Выделяем общие показатели по региону разными цветами
            tags = ()
            if result.get('is_region_total', False):
                category = result.get('category', '')
                if category == '100%':
                    tags = ('perfect_total',)
                elif category == '75-100%':
                    tags = ('good_total',)
            
            self.tree.insert("", "end", values=(
                result['district'],
                score_text,
                result['record_count']
            ), tags=tags)
        
        # Настраиваем стили для строк региона
        self.tree.tag_configure('perfect_total', background='#2e7d32', foreground='white')  # зеленый
        self.tree.tag_configure('good_total', background='#3a7ca5', foreground='white')     # синий
    
    def save_results(self):
        if not self.results:
            messagebox.showwarning("Предупреждение", "Нет данных для сохранения")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Сохранить результаты",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                # Создаем DataFrame с результатами
                df_results = pd.DataFrame(self.results)
                
                # Сохраняем в Excel
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df_results.to_excel(writer, index=False)
                
                # Создаем PDF с диаграммой
                pdf_path = file_path.replace('.xlsx', '_диаграмма.pdf')
                self.create_chart_pdf(pdf_path)
                
                messagebox.showinfo("Успех", "Результаты успешно сохранены!")
                self.status_var.set("Результаты сохранены")
                
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при сохранении: {str(e)}")
    
    def create_chart_pdf(self, pdf_path):
        """Создает PDF с диаграммой"""
        if not self.results:
            return
            
        analysis_type = self.analysis_var.get()
        
        if analysis_type == "score5":
            self.create_bar_chart_pdf(pdf_path)
        else:
            self.create_menu_compliance_chart_pdf(pdf_path)
    
    def create_bar_chart_pdf(self, pdf_path):
        """Создает столбчатую диаграмму для балла нормированного к 5"""
        # Исключаем общие показатели по региону из диаграммы
        chart_data = [r for r in self.results if not r.get('is_region_total', False)]
        
        if not chart_data:
            return
            
        # Подготовка данных
        districts = [r['district'] for r in chart_data]
        scores = [r['score'] for r in chart_data]
        
        # Создание диаграммы с темной темой
        plt.style.use('dark_background')
        fig, ax = plt.subplots(figsize=(12, 8))
        
        colors = ['#3a7ca5' if score == max(scores) else '#5e5e5e' for score in scores]
        
        bars = ax.barh(districts, scores, color=colors, alpha=0.8)
        
        ax.set_xlabel('Показатель (%)', fontsize=12)
        ax.set_title('Балл нормированный к 5 по районам Чеченской республики', fontsize=14, pad=20)
        
        # Добавление значений на диаграмму
        for bar, score in zip(bars, scores):
            width = bar.get_width()
            ax.text(width + 1, bar.get_y() + bar.get_height()/2, 
                   f'{score}%', ha='left', va='center', fontsize=10)
        
        plt.tight_layout()
        fig.savefig(pdf_path, bbox_inches='tight')
        plt.close(fig)
    
    def create_menu_compliance_chart_pdf(self, pdf_path):
        """Создает два круга для соответствия типовому меню"""
        # Получаем общие показатели по региону
        region_data = [r for r in self.results if r.get('is_region_total', False)]
        
        if not region_data:
            return
        
        # Создаем фигуру с тремя subplots
        plt.style.use('dark_background')
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(15, 12))
        
        # Первый круг - 100% соответствие
        perfect_percentage = next((r['score'] for r in region_data if r.get('category') == '100%'), 0)
        perfect_sizes = [perfect_percentage, 100 - perfect_percentage]
        perfect_colors = ['#2e7d32', '#5e5e5e']  # зеленый и серый
        perfect_labels = ['100% соответствие', 'Не 100% соответствие']
        
        wedges1, texts1, autotexts1 = ax1.pie(perfect_sizes, colors=perfect_colors, autopct='%1.1f%%',
                                             startangle=90, shadow=True)
        ax1.set_title('100% соответствие типовому меню', fontsize=14, pad=20)
        ax1.legend(wedges1, perfect_labels, loc="center left", bbox_to_anchor=(0.9, 0, 0.5, 1))
        
        # Второй круг - 75-100% соответствие
        good_percentage = next((r['score'] for r in region_data if r.get('category') == '75-100%'), 0)
        good_sizes = [good_percentage, 100 - good_percentage]
        good_colors = ['#3a7ca5', '#5e5e5e']  # синий и серый
        good_labels = ['75-100% соответствие', 'Менее 75%']
        
        wedges2, texts2, autotexts2 = ax2.pie(good_sizes, colors=good_colors, autopct='%1.1f%%',
                                            startangle=90, shadow=True)
        ax2.set_title('75-100% соответствие типовому меню', fontsize=14, pad=20)
        ax2.legend(wedges2, good_labels, loc="center left", bbox_to_anchor=(0.9, 0, 0.5, 1))
        
        # Третий график - столбчатая диаграмма по районам
        district_data = [r for r in self.results if not r.get('is_region_total', False)]
        
        if district_data:
            districts = [r['district'] for r in district_data]
            scores = [r['score'] for r in district_data]
            
            colors_bar = ['#3a7ca5' if score == max(scores) else '#5e5e5e' for score in scores]
            
            bars = ax3.barh(districts, scores, color=colors_bar, alpha=0.8)
            ax3.set_xlabel('Средний процент соответствия (%)', fontsize=12)
            ax3.set_title('Соответствие по районам', fontsize=14, pad=20)
            
            # Добавление значений на диаграмму
            for bar, score in zip(bars, scores):
                width = bar.get_width()
                ax3.text(width + 1, bar.get_y() + bar.get_height()/2, 
                       f'{score}%', ha='left', va='center', fontsize=9)
        
        # Четвертый график - сводная информация
        ax4.axis('off')
        summary_text = f"""
        СВОДНАЯ ИНФОРМАЦИЯ ПО РЕГИОНУ:
        
        
        
        100% соответствие: {perfect_percentage}%
        75-100% соответствие: {good_percentage}%
        
        Лучший район: {max(district_data, key=lambda x: x['score'])['district'] if district_data else 'Н/Д'}
        Худший район: {min(district_data, key=lambda x: x['score'])['district'] if district_data else 'Н/Д'}
        """
        
        ax4.text(0.1, 0.9, summary_text, fontsize=12, va='top', 
                bbox=dict(boxstyle="round", facecolor='#3d3d3d', alpha=0.7))
        
        plt.tight_layout()
        fig.savefig(pdf_path, bbox_inches='tight')
        plt.close(fig)

def main():
    root = tk.Tk()
    app = DarkDistrictAnalyzerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
