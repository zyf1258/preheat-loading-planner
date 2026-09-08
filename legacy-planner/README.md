# 预热炉工作台（第一期）

纯排炉样板：导入当天生产计划和 MES 棒料登记表后，按现有炉群规则进行逐列排炉。无时间线、炉况控制或设备控制。

## 直接使用软件

双击构建产物：

```text
dist\PreheatFurnaceWorkbench.exe
```

软件自带 Python、Node、前端、Excel 模板和导出依赖，用户无需另外安装运行环境。长度记忆保存在：

```text
%LOCALAPPDATA%\PreheatFurnaceWorkbench
```

推荐使用桌面的“预热炉排炉工作台-快速版”快捷方式。快速版采用目录打包，启动约 2 秒；原来的单文件 exe 仍保留，但每次启动需要解压运行时，启动会明显更慢。

## 开发运行

```powershell
python app.py
```

开发服务地址：`http://127.0.0.1:8760`。

数量统一读取计划表 E 列“数量”；早中夜排产区不参与数量拆分。
