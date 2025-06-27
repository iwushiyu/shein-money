import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re
from datetime import datetime

class CostCalculator:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title('shein工厂货款计算器(wushiyu)')
        self.window.geometry('800x600')
        
        # 创建主框架
        self.main_frame = ttk.Frame(self.window, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建按钮和标签
        self.select_btn = ttk.Button(self.main_frame, text='选择Excel文件', command=self.select_file)
        self.select_btn.pack(pady=10)
        
        self.file_label = ttk.Label(self.main_frame, text='未选择文件')
        self.file_label.pack(pady=5)
        
        self.progress_label = ttk.Label(self.main_frame, text='')
        self.progress_label.pack(pady=5)
        
        self.result_label = ttk.Label(self.main_frame, text='总成本: ¥0.00')
        self.result_label.pack(pady=5)
        
        # 创建预览表格
        self.create_preview_table()
        
        # 保存按钮(初始禁用)
        self.save_btn = ttk.Button(self.main_frame, text='保存结果', command=self.save_results, state='disabled')
        self.save_btn.pack(pady=10)
        
        # 存储计算结果
        self.results_df = None
        
    def create_preview_table(self):
        # 创建预览表格框架
        self.preview_frame = ttk.Frame(self.main_frame)
        self.preview_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建Treeview用于显示预览数据
        columns = ('货号', '材质单价', '面积', '数量', '成本')
        self.preview_table = ttk.Treeview(self.preview_frame, columns=columns, show='headings')
        
        # 设置列标题
        for col in columns:
            self.preview_table.heading(col, text=col)
            self.preview_table.column(col, width=100)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.preview_frame, orient=tk.VERTICAL, command=self.preview_table.yview)
        self.preview_table.configure(yscrollcommand=scrollbar.set)
        
        # 放置表格和滚动条
        self.preview_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
        if file_path:
            self.file_label.config(text=f'已选择: {file_path}')
            try:
                self.calculate_cost(file_path)
                self.save_btn.config(state='normal')
            except Exception as e:
                messagebox.showerror('错误', str(e))
    
    def calculate_cost(self, file_path):
        # 清空预览表格
        for item in self.preview_table.get_children():
            self.preview_table.delete(item)
            
        try:
            # 读取Excel文件
            sales_data = pd.read_excel(file_path, sheet_name=0)
            material_data = pd.read_excel(file_path, sheet_name='材质表', header=None)
            
            # 打印列名用于调试
            print("销售数据表列名:", list(sales_data.columns))
            
            # 检查必要的列是否存在
            required_columns = {
                '商品名称': '货号',  # 使用实际的列名
                '规格': '属性集',
                '数量': '下单数量'  # 需要确认数量列的名称
            }
            
            # 检查是否找到所有必要的列
            missing_columns = [key for key, value in required_columns.items() if value is None]
            if missing_columns:
                raise ValueError(f"找不到必要的列: {', '.join(missing_columns)}\n实际列名: {list(sales_data.columns)}")
            
            # 创建材质价格字典,并按材质名称长度降序排序
            material_items = sorted(zip(material_data[0], material_data[1]), 
                                  key=lambda x: len(str(x[0])), reverse=True)
            material_prices = dict(material_items)
            
            # 创建结果DataFrame
            results = []
            total_cost = 0
            processed_rows = 0
            
            for idx, row in sales_data.iterrows():
                try:
                    # 提取货号
                    product_name = str(row[required_columns['商品名称']])
                    
                    # 查找匹配的材质
                    found_material = None
                    for material in material_prices.keys():
                        if str(material) in product_name:
                            found_material = material
                            break
                            
                    if not found_material:
                        print(f"行 {idx+2} 找不到匹配的材质: {product_name}")
                        continue
                        
                    material_price = material_prices[found_material]
                    
                    # 计算面积
                    spec = str(row[required_columns['规格']])
                    try:
                        if '直径' in spec:
                            diameter = float(re.search(r'直径(\d+)', spec).group(1))
                            area = (diameter/100) ** 2
                        else:
                            # 提取规格中的数字
                            numbers = re.findall(r'\d+', spec)
                            if len(numbers) >= 2:
                                length, width = map(float, numbers[:2])
                                area = length * width * 0.0001
                            else:
                                raise ValueError(f"无法从规格中提取尺寸: {spec}")
                    except (ValueError, AttributeError) as e:
                        print(f"行 {idx+2} 规格格式错误: {spec}")
                        continue
                    
                    # 获取数量 (需要确认数量列的名称)
                    try:
                        quantity = float(row[required_columns['数量']])
                    except (ValueError, TypeError):
                        print(f"行 {idx+2} 数量格式错误: {row[required_columns['数量']]}")
                        quantity = 0  # 或者 continue，取决于您希望如何处理错误数据
                    
                    # 计算成本
                    cost = material_price * area * quantity
                    total_cost += cost
                    
                    # 添加到结果列表
                    results.append({
                        '货号': product_name,
                        '材质单价': f'{material_price:.2f}', #如果要单位可以 f'¥{material_price:.2f}'
                        '面积': f'{area:.4f}㎡',
                        '数量': quantity,
                        '成本': f'{cost:.2f}'  #如果要单位可以 f'¥{material_price:.2f}'
                    })
                    
                    # 更新预览表格
                    self.preview_table.insert('', tk.END, values=(
                        product_name,
                        f'{material_price:.2f}',
                        f'{area:.4f}㎡',
                        quantity,
                        f'{cost:.2f}'
                    ))
                    
                    processed_rows += 1
                    self.progress_label.config(text=f'已处理 {processed_rows} 行')
                    self.window.update()
                    
                except Exception as e:
                    print(f'处理行 {idx+2} 时出错: {str(e)}')
                    continue
            
            # 保存结果DataFrame
            self.results_df = pd.DataFrame(results)
            
            # 更新总成本显示
            self.result_label.config(text=f'总成本: ¥{total_cost:.2f}')
            
            if processed_rows == 0:
                messagebox.showwarning('警告', '没有成功处理任何数据行')
            else:
                messagebox.showinfo('完成', f'成功处理 {processed_rows} 行数据')
                
        except Exception as e:
            raise Exception(f"处理文件时出错: {str(e)}")
    
    def save_results(self):
        if self.results_df is None or len(self.results_df) == 0:
            messagebox.showwarning('警告', '没有可保存的结果')
            return
            
        # 获取保存路径
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        default_filename = f'成本计算结果_{timestamp}.xlsx'
        save_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[('Excel Files', '*.xlsx')],
            initialfile=default_filename
        )
        
        if save_path:
            try:
                # 保存结果
                self.results_df.to_excel(save_path, index=False)
                messagebox.showinfo('成功', '结果已保存')
            except Exception as e:
                messagebox.showerror('错误', f'保存失败: {str(e)}')
    
    def run(self):
        self.window.mainloop()

if __name__ == '__main__':
    app = CostCalculator()
    app.run()