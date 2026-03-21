# 26年NC台账管理系统

## 快速启动

在终端运行以下命令启动服务：

```bash
bash ~/.qclaw/workspace/nc_account_system/start_server.sh
```

首次运行后会显示外网访问地址，类似：
```
https://xxxxxx.serveousercontent.com
```

## 访问方式

- **外网访问**：上面的 https://xxxxx 地址（任何设备打开即用）
- **局域网访问**：http://192.168.0.155:5001（需同一WiFi）

## 功能说明

1. **查看数据** - 首页仪表盘展示统计和图表
2. **搜索** - 可按包裹号、商品详情等搜索
3. **添加数据** - 点击「添加新数据」按钮
4. **编辑/删除** - 每行数据后面有操作按钮
5. **导出Excel** - 点击「导出Excel」下载完整数据

## 数据存储

所有数据会实时同步保存到桌面原始Excel文件：
```
~/Desktop/26年NC台账勿删.xlsx
```

---
如遇外网链接失效，重新运行启动脚本即可获得新链接。
# Render deploy trigger Sat Mar 21 19:27:04 CST 2026
