import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import wmi


class IPConfigTool:
    def __init__(self, root):
        self.root = root
        self.root.title("网卡IP快捷修改工具")
        
        self.config_file = os.path.join(os.getcwd(), "ipedit_config.json")
        
        # 先加载配置，设置窗口大小
        self.favorites = self.load_config()
        
        self.setup_ui()
        self.refresh_network_adapters()
        
        self.root.bind("<Configure>", self.save_window_size)
        self.paned_window.bind("<ButtonRelease-1>", self.save_sash_position)
        self.paned_window.bind("<B1-Motion>", self.save_sash_position)
        
        # 窗口大小已在 load_config 中设置
        self.root.update_idletasks()  # 确保窗口大小已应用
        self.load_sash_position()
        
    def load_window_size(self):
        # 窗口大小由 load_config 统一加载
        pass
    
    def save_window_size(self, event=None):
        # 窗口大小改变时保存配置
        self.save_config()
    
    def save_sash_position(self, event=None):
        # 分隔条位置改变时保存配置
        self.save_config()
    
    def load_sash_position(self):
        try:
            if hasattr(self, 'paned_window'):
                # 使用从配置中加载的分隔条位置
                position = getattr(self, 'sash_position', 400)
                # 确保窗口大小已完全应用
                self.root.update_idletasks()
                # 延迟设置分隔条位置，确保窗口布局已完成
                self.root.after(100, lambda: self._apply_sash_position(position))
        except Exception as e:
            print(f"加载分隔条位置失败: {e}")
    
    def _apply_sash_position(self, position):
        try:
            if hasattr(self, 'paned_window'):
                # tk.PanedWindow 使用 sash_place 设置位置
                self.paned_window.sash_place(0, position, 0)
        except Exception as e:
            print(f"应用分隔条位置失败: {e}")
    
    def load_config(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # 加载收藏夹
                    self.favorites = config.get('favorites', [])
                    # 加载窗口大小
                    geometry = config.get('window', {}).get('geometry', '800x400')
                    self.root.geometry(geometry)
                    # 加载分隔条位置
                    self.sash_position = config.get('panel', {}).get('position', 400)
                    # 设置窗口最小大小
                    self.root.minsize(50, 50)
                    return self.favorites
        except Exception as e:
            print(f"加载配置失败: {e}")
        # 默认值
        self.favorites = []
        self.sash_position = 400
        self.root.geometry("800x400")
        self.root.minsize(50, 50)
        return self.favorites
    
    def save_config(self):
        try:
            # 构建完整配置
            config = {
                'favorites': self.favorites,
                'window': {
                    'geometry': self.root.geometry()
                },
                'panel': {
                    'position': self.paned_window.sash_coord(0)[0] if hasattr(self, 'paned_window') else 400
                }
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {e}")
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.paned_window = tk.PanedWindow(main_frame, orient=tk.HORIZONTAL, sashwidth=10, sashrelief='raised')
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        left_frame = ttk.LabelFrame(self.paned_window, text="网卡配置")
        self.paned_window.add(left_frame)
        
        right_frame = ttk.LabelFrame(self.paned_window, text="配置收藏夹")
        self.paned_window.add(right_frame)
        
        self.setup_left_panel(left_frame)
        self.setup_right_panel(right_frame)
        
    def setup_left_panel(self, parent):
        form_frame = ttk.Frame(parent)
        form_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(form_frame, text="网卡名称:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.adapter_var = tk.StringVar()
        self.adapter_combo = ttk.Combobox(form_frame, textvariable=self.adapter_var, state="readonly", width=50)
        self.adapter_combo.grid(row=0, column=1, sticky=tk.EW, pady=5)
        self.adapter_combo.bind("<<ComboboxSelected>>", lambda e: self.load_adapter_config())
        
        ttk.Label(form_frame, text="IP地址:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.ip_entry = ttk.Entry(form_frame, width=30)
        self.ip_entry.grid(row=1, column=1, sticky=tk.EW, pady=5)
        
        ttk.Label(form_frame, text="子网掩码:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.subnet_entry = ttk.Entry(form_frame, width=30)
        self.subnet_entry.grid(row=2, column=1, sticky=tk.EW, pady=5)
        
        ttk.Label(form_frame, text="默认网关:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.gateway_entry = ttk.Entry(form_frame, width=30)
        self.gateway_entry.grid(row=3, column=1, sticky=tk.EW, pady=5)
        
        ttk.Label(form_frame, text="DNS服务器:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.dns_entry = ttk.Entry(form_frame, width=30)
        self.dns_entry.grid(row=4, column=1, sticky=tk.EW, pady=5)
        
        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="获取DHCP", command=self.get_dhcp).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="刷新", command=self.load_adapter_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="应用配置", command=self.apply_config).pack(side=tk.LEFT, padx=5)
        
        form_frame.columnconfigure(1, weight=1)
        
    def setup_right_panel(self, parent):
        toolbar_frame = ttk.Frame(parent)
        toolbar_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(toolbar_frame, text="添加到收藏夹", command=self.save_current_to_favorites).pack(side=tk.LEFT, padx=2)
        
        list_frame = ttk.Frame(parent)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.favorites_listbox = tk.Listbox(list_frame, height=8)
        self.favorites_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.favorites_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.favorites_listbox.config(yscrollcommand=scrollbar.set)
        
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(button_frame, text="← 加载配置", command=self.load_favorite_config).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="删除选中", command=self.delete_favorite).pack(side=tk.LEFT, padx=2)
        
        self.refresh_favorites()
        
    def refresh_network_adapters(self):
        try:
            c = wmi.WMI()
            adapters = []
            for adapter in c.Win32_NetworkAdapterConfiguration(IPEnabled=True):
                adapters.append(adapter.Description)
            self.adapter_combo['values'] = adapters
            if adapters:
                self.adapter_combo.current(0)
                self.load_adapter_config()
        except Exception as e:
            messagebox.showerror("错误", f"获取网卡信息失败: {e}")
    
    def load_adapter_config(self):
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            return
        
        try:
            c = wmi.WMI()
            for adapter in c.Win32_NetworkAdapterConfiguration(IPEnabled=True):
                if adapter.Description == adapter_name:
                    # 清空所有输入框
                    self.ip_entry.delete(0, tk.END)
                    self.subnet_entry.delete(0, tk.END)
                    self.gateway_entry.delete(0, tk.END)
                    self.dns_entry.delete(0, tk.END)
                    
                    # 只在有值时填充
                    if adapter.IPAddress:
                        self.ip_entry.insert(0, adapter.IPAddress[0])
                    if adapter.IPSubnet:
                        self.subnet_entry.insert(0, adapter.IPSubnet[0])
                    if adapter.DefaultIPGateway:
                        # 确保DefaultIPGateway不为空且包含有效值
                        gateways = [g for g in adapter.DefaultIPGateway if g]
                        if gateways:
                            self.gateway_entry.insert(0, ",".join(gateways))
                    if adapter.DNSServerSearchOrder:
                        # 确保DNSServerSearchOrder不为空且包含有效值
                        dns_servers = [d for d in adapter.DNSServerSearchOrder if d]
                        if dns_servers:
                            self.dns_entry.insert(0, ",".join(dns_servers))
                    break
        except Exception as e:
            messagebox.showerror("错误", f"加载网卡配置失败: {e}")
    
    def apply_config(self):
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            messagebox.showwarning("警告", "请先选择网卡")
            return
        
        ip = self.ip_entry.get().strip()
        subnet = self.subnet_entry.get().strip()
        gateway = self.gateway_entry.get().strip()
        dns = self.dns_entry.get().strip()
        
        try:
            c = wmi.WMI()
            for adapter in c.Win32_NetworkAdapterConfiguration(IPEnabled=True):
                if adapter.Description == adapter_name:
                    if ip and subnet:
                        net_mask = [subnet]
                        result = adapter.EnableStatic(IPAddress=[ip], SubnetMask=net_mask)
                        if result[0] != 0:
                            messagebox.showerror("错误", f"设置IP失败，返回码: {result[0]}")
                            return
                    
                    if gateway:
                        gateway_list = [g.strip() for g in gateway.split(",") if g.strip()]
                        result = adapter.SetDefaultGateway(DefaultIPGateway=gateway_list)
                        if result[0] != 0:
                            messagebox.showerror("错误", f"设置网关失败，返回码: {result[0]}")
                            return
                    
                    if dns:
                        dns_list = [d.strip() for d in dns.split(",") if d.strip()]
                        result = adapter.SetDNSServerSearchOrder(DNSServerSearchOrder=dns_list)
                        if result[0] != 0:
                            messagebox.showerror("错误", f"设置DNS失败，返回码: {result[0]}")
                            return
                    
                    messagebox.showinfo("成功", "配置已应用")
                    break
        except Exception as e:
            messagebox.showerror("错误", f"应用配置失败: {e}")
    
    def get_dhcp(self):
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            messagebox.showwarning("警告", "请先选择网卡")
            return
        
        try:
            c = wmi.WMI()
            for adapter in c.Win32_NetworkAdapterConfiguration(IPEnabled=True):
                if adapter.Description == adapter_name:
                    adapter.EnableDHCP()
                    messagebox.showinfo("成功", "已获取DHCP配置")
                    self.load_adapter_config()
                    break
        except Exception as e:
            messagebox.showerror("错误", f"获取DHCP失败: {e}")
    
    def save_current_to_favorites(self):
        ip = self.ip_entry.get().strip()
        name = ip if ip else "未命名配置"
        
        config = {
            "name": name,
            "ip": ip,
            "subnet": self.subnet_entry.get().strip(),
            "gateway": self.gateway_entry.get().strip(),
            "dns": self.dns_entry.get().strip()
        }
        
        self.favorites.append(config)
        self.save_config()
        self.refresh_favorites()
    
    def refresh_favorites(self):
        self.favorites_listbox.delete(0, tk.END)
        for fav in self.favorites:
            self.favorites_listbox.insert(tk.END, fav['name'])
    
    def load_favorite_config(self):
        selection = self.favorites_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一个收藏项")
            return
        
        index = selection[0]
        config = self.favorites[index]
        
        self.ip_entry.delete(0, tk.END)
        self.ip_entry.insert(0, config['ip'])
        self.subnet_entry.delete(0, tk.END)
        self.subnet_entry.insert(0, config['subnet'])
        self.gateway_entry.delete(0, tk.END)
        self.gateway_entry.insert(0, config['gateway'])
        self.dns_entry.delete(0, tk.END)
        self.dns_entry.insert(0, config['dns'])
    
    def delete_favorite(self):
        selection = self.favorites_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要删除的收藏项")
            return
        
        index = selection[0]
        if messagebox.askyesno("确认删除", "确定要删除这个收藏项吗？"):
            del self.favorites[index]
            self.save_config()
            self.refresh_favorites()


def main():
    root = tk.Tk()
    app = IPConfigTool(root)
    root.mainloop()


if __name__ == "__main__":
    main()

