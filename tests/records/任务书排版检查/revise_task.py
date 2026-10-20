from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from copy import deepcopy
from lxml import etree
import shutil, sys, html

root = Path(__file__).resolve().parents[3]
target = next((root/'docs/00-资料/任务书').glob('*.docx'))
archive = root/'archive/任务书修订/2026-10-13'
backup = archive/target.name
if not backup.exists():
    shutil.copy2(target, backup)
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ns = {'w': W}
q = lambda name: '{'+W+'}'+name
goals = [
'本课题以西门子S7-1200 PLC为主控制器，面向单泵恒压供水实验装置，采用低压直流水泵、压力传感器及调速驱动器构建压力闭环。通过改变回水阀开度模拟用水量变化，研究PLC在负载扰动下的压力调节与故障保护。具体任务如下：',
'1. 完成总体方案及器件选型。以CPU 1211C DC/DC/DC为基础，设计水箱—水泵—压力测量—用水调节阀—回流水箱的循环水路，核对传感器量程、水泵扬程、驱动接口及电源容量，编制物料表、I/O分配表和接线图。',
'2. 参考已有液位PID工程，在TIA Portal中理解模拟量处理、PID及HMI程序结构，完成压力变量、调速输出和报警逻辑的适配；先进行PLCSIM程序验证，再开展实物联调，记录开源参考与自主修改的内容。',
'3. 实现压力信号采集、标定和工程量换算，建立设定压力、实际压力和调速输出之间的闭环关系；采用PLC内部PID控制，根据压力偏差调节水泵转速，并完成参数整定、输出限幅及启动斜坡设计。',
'4. 实现手动／自动运行、启停、故障复位和参数设置功能，设计压力趋势、运行状态及报警显示界面，验证不同模式下的操作权限和保护优先级。',
'5. 完成设定值变化、用水量扰动、缺水及传感器异常等测试，分析稳态误差、超调量、调节时间和压力波动；形成可复核的测试数据、程序工程、设计图纸及毕业论文。'
]
requirements = [
'☑有  □无   作品设计要求：',
'1. 装置与接口。制作单泵低压循环水路，配置压力传感器、机械压力表、回水调节阀及低水位检测装置。压力输入优先采用0～10 V信号接入PLC模拟量输入；采用SB 1232模拟量输出板或经核验的等效接口，向匹配的调速驱动器提供给定信号，由驱动器完成直流水泵调速。PLC输出不得直接承担水泵功率驱动。',
'2. 压力采集。按照模块配置核对模拟量原始值范围，通过NORM_X、SCALE_X或等效计算换算为MPa等统一压力单位；完成零点、量程及机械压力表对照标定。对信号进行适当滤波，设置越界和异常变化判据，保留零压力正常启动的处理逻辑。',
'3. PID控制。采用PID_Compact或适配的PID程序，在固定周期中断中执行运算，采样周期可从100 ms起进行调试，保持控制块周期参数与调用周期一致。核对控制方向，使压力低于设定值时增加有效调速输出；设置输出上下限、积分饱和处理及启停复位，避免停机后积分累积造成再次启动冲击。',
'4. 操作逻辑。自动模式根据压力反馈调速，手动模式允许在限定范围内调整输出；两种模式均受急停和故障联锁约束。切换时采用输出跟踪或过渡处理，限制调速给定突变。压力设定值和报警阈值应限制在器件及水路允许范围内。停电恢复或故障消除后须经人工确认启动。',
'5. 保护与报警。至少设置缺水、超压、传感器异常及驱动器故障的停泵联锁；过流／过载保护由具备该功能的驱动器或独立保护器件承担。保护触发时关闭运行使能并将调速给定置零，显示故障原因；故障条件未解除时禁止复位启动。欠压报警结合启动延时和持续时间判断，避免正常建压阶段误报警。',
'6. 监控与记录。在TIA中设计可仿真的HMI界面，显示设定压力、实际压力、压力偏差、调速输出、运行模式和报警状态，支持必要参数设置及压力趋势观察；实物阶段可采用电脑监控，无须另购TP700面板。记录测试时间、设定值、测量值、输出及报警状态，便于整理为CSV数据。',
'7. 测试与评价。在安全工作范围内选取至少三个压力设定点，进行建压、设定值阶跃及阀门开度扰动测试，关键工况至少重复三次；按统一方法计算稳态误差、最大超调和调节时间，并保留曲线及原始记录。通过模拟故障验证停泵、报警和复位联锁。压力范围及验收阈值应依据最终选型、标定结果和指导教师要求确定。',
'8. 装配与交付。采用适配的低压直流电源，设置分路保险、急停切断水泵使能或动力回路，并采取干湿分区、管路防脱和必要泄压措施。完成端子编号、线缆标识及设备固定，提交实物装置、PLC/HMI工程、接线图、物料清单、操作说明、测试报告和论文。明确本装置采用直流水泵调速，不将其表述为交流变频器供水系统。'
]
purpose = [
'本课题通过单泵循环水路建立可观察、可调试的恒压供水实验平台，将PLC模拟量采集、PID闭环控制、人机监控及故障联锁应用于压力调节。利用用水量扰动与设定值变化测试，分析控制参数对稳态误差、响应速度和压力波动的影响，掌握从器件选型、接线编程到系统联调的完整设计方法。采用低压直流水泵有利于控制实验成本，适合教学和毕业设计验证；成果用于说明恒压控制原理与设计方法，应用于实际供水场景还需进一步完成工程选型和可靠性验证。'
]
replacements={1:['基于PLC的恒压供水控制系统设计'],27:goals,33:requirements,35:purpose}
with ZipFile(backup) as z:
    original = z.read('word/document.xml')
    doc = etree.fromstring(original)
    cells = doc.xpath('//w:tc',namespaces=ns)
    old_text={i:''.join(tc.xpath('.//w:t/text()',namespaces=ns)) for i,tc in enumerate(cells)}
    for index,texts in replacements.items():
        cell=cells[index]
        old_p=cell.find('w:p',ns)
        ppr=deepcopy(old_p.find('w:pPr',ns))
        for p in list(cell.findall('w:p',ns)):cell.remove(p)
        for text in texts:
            p=etree.SubElement(cell,q('p'))
            if ppr is not None:
                props=deepcopy(ppr)
                for name in ['rPr','ind','keepNext','pageBreakBefore']:
                    for item in props.findall('w:'+name,ns):props.remove(item)
                spacing=props.find('w:spacing',ns)
                if spacing is None:spacing=etree.SubElement(props,q('spacing'))
                spacing.set(q('line'),'270');spacing.set(q('lineRule'),'auto')
                spacing.set(q('before'),'0');spacing.set(q('after'),'35')
                p.append(props)
            r=etree.SubElement(p,q('r'));rp=etree.SubElement(r,q('rPr'))
            fonts=etree.SubElement(rp,q('rFonts'))
            for name in ['ascii','hAnsi','eastAsia','cs']:fonts.set(q(name),'宋体')
            etree.SubElement(rp,q('sz')).set(q('val'),'21')
            etree.SubElement(rp,q('szCs')).set(q('val'),'21')
            etree.SubElement(rp,q('color')).set(q('val'),'000000')
            etree.SubElement(r,q('t')).text=text
        if index in [27,33,35]:
            row=cell.getparent();trpr=row.find('w:trPr',ns)
            if trpr is not None:
                for item in list(trpr):
                    if etree.QName(item).localname in ['trHeight','cantSplit']:trpr.remove(item)
    # The template has a blank continuation row below the task-goal cell.
    rows=doc.xpath('//w:tr',namespaces=ns)
    for item in rows[7].xpath('./w:trPr/w:trHeight',namespaces=ns):item.set(q('val'),'80');item.set(q('hRule'),'atLeast')
    output=etree.tostring(doc,xml_declaration=True,encoding='UTF-8',standalone=True)
    staging=target.with_suffix('.revised.docx')
    with ZipFile(staging,'w',ZIP_DEFLATED) as out:
        for item in z.infolist():out.writestr(item,output if item.filename=='word/document.xml' else z.read(item.filename))
    staging.replace(target)

# Preserve a tracked textual review copy with the Document library.
unpacked=archive/'review-unpacked'
with ZipFile(backup) as z:z.extractall(unpacked)
sys.path.insert(0,r'C:\Users\m2714\.codex\skills\docx')
sys.path.insert(0,str(Path(__file__).parent/'python_deps'))
from scripts.document import Document
review=Document(unpacked,author='Codex',initials='CX')
editor=review['word/document.xml']
for index,texts in replacements.items():
    node=editor.get_node(tag='w:p',contains=old_text[index])
    props=node.getElementsByTagName('w:pPr')
    prop_xml=props[0].toxml() if props else ''
    deleted=html.escape(old_text[index])
    inserted='<w:br/>'.join('<w:t>'+html.escape(t)+'</w:t>' for t in texts)
    editor.replace_node(node,f'<w:p>{prop_xml}<w:del><w:r><w:delText>{deleted}</w:delText></w:r></w:del><w:ins><w:r>{inserted}</w:r></w:ins></w:p>')
review.save()
with ZipFile(archive/'任务书_文字修订留痕.docx','w',ZIP_DEFLATED) as z:
    for p in unpacked.rglob('*'):
        if p.is_file():z.write(p,p.relative_to(unpacked).as_posix())
summary=root/'docs/01-立项/任务书修订说明.md'
summary.write_text('# 任务书修订说明\n\n日期：2026-10-13。\n\n已就地修改任务书的题目、课题主要任务、作品设计要求和选题意义；姓名、学号、教师、教学安排、签字栏及原日期保持不变。\n\n按此前讨论的西门子单泵低压循环装置路线，将交流变频调速改为直流水泵调速，补充模拟量换算、PID周期控制、调速驱动接口、模式切换、保护联锁和测试方法。最终器件型号、压力范围及验收阈值需结合选型和教师要求确定；本次文本不作为功能已实现的证明。\n\n原文件及文字修订留痕保存于 `archive/任务书修订/2026-10-13/`。之前的文件整理校验清单是整理当日基线，当前任务书经授权修改后哈希发生变化。\n',encoding='utf-8')
print('Revised:',target)
