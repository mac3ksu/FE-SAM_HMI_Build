import xlrd


def build_pts_dict(worksheet_obj):
    i = 1
    pts_dict = {}
    while i < worksheet_obj.nrows:
        dict_key = int(worksheet_obj.cell_value(i, 0))
        dnp_address = dict_key
        desc = str(worksheet_obj.cell_value(i, 1)).upper()
        desc = desc.replace('"', '')
        desc = desc.replace('.', '')
        desc = desc.replace('–', '-')
        desc = desc.replace('&', 'and')
        # print(i, desc)
        state_0_desc = worksheet_obj.cell_value(i, 2)
        state_1_desc = worksheet_obj.cell_value(i, 3)
        alarm_state = int(worksheet_obj.cell_value(i, 4))

        pts_dict[dict_key] = (dnp_address, desc, state_0_desc, state_1_desc, alarm_state)
        i += 1
    #print(pts_dict)
    return pts_dict


def build_page_list(worksheet_obj, row):
    page_list = [
                    worksheet_obj.cell_value(row, 6),
                    worksheet_obj.cell_value(row, 7),
                    worksheet_obj.cell_value(row, 8),
                    worksheet_obj.cell_value(row, 9),
                    worksheet_obj.cell_value(row, 10),
                    worksheet_obj.cell_value(row, 11),
                    worksheet_obj.cell_value(row, 12),
                    worksheet_obj.cell_value(row, 13),
                    worksheet_obj.cell_value(row, 14),
                    worksheet_obj.cell_value(row, 15),
                    worksheet_obj.cell_value(row, 16),
                    worksheet_obj.cell_value(row, 17),
                    worksheet_obj.cell_value(row, 18),
                    worksheet_obj.cell_value(row, 19),
                    worksheet_obj.cell_value(row, 20),
                    worksheet_obj.cell_value(row, 21),
                    worksheet_obj.cell_value(row, 22),
                    worksheet_obj.cell_value(row, 23),
                    worksheet_obj.cell_value(row, 24),
                    worksheet_obj.cell_value(row, 25),
                    worksheet_obj.cell_value(row, 26),
                    worksheet_obj.cell_value(row, 27),
                    worksheet_obj.cell_value(row, 28),
                    worksheet_obj.cell_value(row, 29),
                    worksheet_obj.cell_value(row, 30),
                    worksheet_obj.cell_value(row, 31),
                ]
    for i, item in enumerate(page_list):
        try:
            page_list[i] = int(item)
        except:
            page_list[i] = item
    return page_list


def print_hmi_point(output, mapped_pts_list, pts_dict):
    for pt in mapped_pts_list:

        dnp_add = pts_dict[pt][0]
        name = pts_dict[pt][1]
        state_0 = pts_dict[pt][2]
        state_1 = pts_dict[pt][3]
        alm_state = pts_dict[pt][4]

        name = name.replace('"', '')
        name = name.replace('.', '')

        if name != '':
            alm_state = int(alm_state)
            output.write('          <Item Name="DI_{}" Path="" DisplayName="{}" DataType="1" AccessRight="1" PointType="1" AudibleDisable="False" RemoveFromTabular="False">\n'.format(
                            str(dnp_add),
                            str(name))
                        )
            output.write('            <TransitionState>0</TransitionState>\n')
            output.write('            <TransitionTimeout>0</TransitionTimeout>\n')
            output.write('            <ControlTimeout>0</ControlTimeout>\n')
            if alm_state:
                output.write('            <State1Abnormal>True</State1Abnormal>\n')
                output.write('            <State0Abnormal>False</State0Abnormal>\n')
            else:
                output.write('            <State1Abnormal>False</State1Abnormal>\n')
                output.write('            <State0Abnormal>True</State0Abnormal>\n')
            if state_0 == '':
                output.write('            <State1Text />\n')
                output.write('            <State0Text />\n')
            else:
                output.write('            <State1Text>{}</State1Text>\n'.format(str(state_1)))
                output.write('            <State0Text>{}</State0Text>\n'.format(str(state_0)))
            output.write('            <RemoveFromTabular>False</RemoveFromTabular>\n')
            output.write('          </Item>\n')
        else:
            output.write('          <Item Name="DI_{}" Path="" DisplayName="Point_{}" DataType="1" AccessRight="1" PointType="1" AudibleDisable="False" RemoveFromTabular="False">\n'.format(
                    str(dnp_add),
                    str(dnp_add))
            )
            output.write('            <TransitionState>0</TransitionState>\n')
            output.write('            <TransitionTimeout>0</TransitionTimeout>\n')
            output.write('            <ControlTimeout>0</ControlTimeout>\n')
            output.write('            <State1Abnormal>False</State1Abnormal>\n')
            output.write('            <State0Abnormal>True</State0Abnormal>\n')
            output.write('            <State1Text />\n')
            output.write('            <State0Text />\n')
            output.write('            <RemoveFromTabular>False</RemoveFromTabular>\n')
            output.write('          </Item>\n')


def build_pages(output, pages_list, pts_dict):
    for page in pages_list:
        ints = 0
        for item in page:
            if isinstance(item, int):
                ints += 1
        if ints > 0:
            output.write('        <Page Name="{}">\n'.format(page[0]))
            x = 0
            y = 0
            for alm in page[1:]:
                if isinstance(alm, int):
                    output_file.write('          <Button Name="{}" X="{}" Y="{}" Width="1" Height="1">\n'.format(pts_dict[alm][1], x, y))
                    output_file.write('            <Action>12</Action>\n')
                    output_file.write('            <Link />\n')
                    output_file.write('            <Endpoint />\n')
                    output_file.write('            <Equipment>Device_0</Equipment>\n')
                    output_file.write('            <ItemName>DI_{}</ItemName>\n'.format(pts_dict[alm][0]))
                    output_file.write('            <FeedbackName />\n')
                    output_file.write('            <Modifier>0</Modifier>\n')
                    output_file.write('          </Button>\n')

                    x += 1
                    if x > 4:
                        x = 0
                        y += 1
                else:
                    x += 1
                    if x > 4:
                        x = 0
                        y += 1
            output_file.write('        </Page>\n')
        else:
            output_file.write('        <Page Name="{}" />\n'.format(page[0]))

if __name__ == '__main__':
    sub_name = 'WHIPPANY SUB'
    outfile = 'Whippany SAM Rev A.SAM'

    wbook = xlrd.open_workbook('Whippany SAM HMI Build.xlsx')
    wsheet = wbook.sheet_by_index(0)

    number_of_pages = -1

    for cell in wsheet.col(6):
        if cell.value != '':
            number_of_pages += 1

    annun_pts_list = []
    pts_dict = build_pts_dict(wsheet)
    mapped_pts_list = sorted(pts_dict.keys())

    # print('Pages:')
    pages_list = []
    for row in range(1,1+number_of_pages):
        # print(row)
        pages_list.append(build_page_list(wsheet, row))

    for page in pages_list:
        for item in page:
            if isinstance(item, int):
                annun_pts_list.append(item)

    print('The following DNP points arent mapped to a page:')
    for pt in mapped_pts_list:
        if pt not in annun_pts_list:
            print(pts_dict[pt][0], pts_dict[pt][1])

    with open(outfile, 'w') as output_file:
        output_file.write('<AlarmConfiguration>\n')
        output_file.write('  <OPCSection Hierarchical="0">\n')
        output_file.write('    <Endpoint Path="http://localhost:8080/AseOpcServer.1" Name="">\n')
        output_file.write('      <Group Name="LINE_0" Path="">\n')
        output_file.write('        <Group Name="Device_0" Path="">\n')

        print_hmi_point(output_file, mapped_pts_list, pts_dict)

        output_file.write('        </Group>\n')
        output_file.write('      </Group>\n')
        output_file.write('    </Endpoint>\n')
        output_file.write('  </OPCSection>\n')
        output_file.write('  <DisplaySection HeaderRowCount="2" HeaderColumnCount="8" RowCount="5" ColumnCount="5" PointSize="12" State0Text="" State1Text="" Vertical="0" HeaderSize="153" FlatAddressing="0" DisplayRTUName="0" DisplayEndpointName="0" AcknowledgeOnReturnToNormal="1" AcknowledgeAllStates="0" Name="{}" Home="MAIN PAGE" SoundFile="c:\windows\media\chimes.wav" AckLog="1" FullScreen="0" SoundTimeout="0" DecimalDigits="0" EventCount="500" ArchiveFile="" GpioAudible="0" AlarmReturnToNormal="1" LongTimeFormat="dd/MM/yyyy HH:mm:ss.FFF" ShortTimeFormat="" PollFrequency="0" LastModified="3/21/2017 5:47:13 PM" ModifiedWith="1.1.1.15">\n'.format(sub_name))
        output_file.write('    <Display Name="Display" X="0" Y="0" Width="0" Height="0">\n')
        output_file.write('      <Menu Name="Menu" X="0" Y="0" Width="8" Height="1">\n')
        output_file.write('        <Button Name="Alarm Summary" X="0" Y="0" Width="1" Height="1">\n')
        output_file.write('          <Action>3</Action>\n')
        output_file.write('          <Link />\n')
        output_file.write('          <Endpoint />\n')
        output_file.write('          <Equipment />\n')
        output_file.write('          <ItemName />\n')
        output_file.write('          <FeedbackName />\n')
        output_file.write('          <Modifier>0</Modifier>\n')
        output_file.write('        </Button>\n')
        output_file.write('        <Button Name="Forward" InitialName="First" X="6" Y="0" Width="1" Height="1">\n')
        output_file.write('          <Action>6</Action>\n')
        output_file.write('          <Link />\n')
        output_file.write('          <Endpoint />\n')
        output_file.write('          <Equipment />\n')
        output_file.write('          <ItemName />\n')
        output_file.write('          <FeedbackName />\n')
        output_file.write('          <Modifier>0</Modifier>\n')
        output_file.write('        </Button>\n')
        output_file.write('        <Button Name="Back" InitialName="Current" X="7" Y="0" Width="1" Height="1">\n')
        output_file.write('          <Action>7</Action>\n')
        output_file.write('          <Link />\n')
        output_file.write('          <Endpoint />\n')
        output_file.write('          <Equipment />\n')
        output_file.write('          <ItemName />\n')
        output_file.write('          <FeedbackName />\n')
        output_file.write('          <Modifier>0</Modifier>\n')
        output_file.write('        </Button>\n')
        output_file.write('        <Button Name="Event Log" X="1" Y="0" Width="1" Height="1">\n')
        output_file.write('          <Action>16</Action>\n')
        output_file.write('          <Link />\n')
        output_file.write('          <Endpoint />\n')
        output_file.write('          <Equipment>Device_0</Equipment>\n')
        output_file.write('          <ItemName>D7 230KV BKR</ItemName>\n')
        output_file.write('          <FeedbackName />\n')
        output_file.write('          <Modifier>0</Modifier>\n')
        output_file.write('        </Button>\n')
        output_file.write('        <Button Name="Silence" X="5" Y="0" Width="1" Height="1">\n')
        output_file.write('          <Action>15</Action>\n')
        output_file.write('          <Link />\n')
        output_file.write('          <Endpoint />\n')
        output_file.write('          <Equipment>Device_0</Equipment>\n')
        output_file.write('          <ItemName>D8 230KV BKR</ItemName>\n')
        output_file.write('          <FeedbackName />\n')
        output_file.write('          <Modifier>0</Modifier>\n')
        output_file.write('        </Button>\n')
        output_file.write('        <Button Name="Tabular" X="2" Y="0" Width="1" Height="1">\n')
        output_file.write('          <Action>14</Action>\n')
        output_file.write('          <Link />\n')
        output_file.write('          <Endpoint />\n')
        output_file.write('          <Equipment />\n')
        output_file.write('          <ItemName />\n')
        output_file.write('          <FeedbackName />\n')
        output_file.write('          <Modifier>0</Modifier>\n')
        output_file.write('        </Button>\n')
        output_file.write('        <Button Name="AckAll" X="4" Y="0" Width="1" Height="1">\n')
        output_file.write('          <Action>19</Action>\n')
        output_file.write('          <Link />\n')
        output_file.write('          <Endpoint />\n')
        output_file.write('          <Equipment />\n')
        output_file.write('          <ItemName />\n')
        output_file.write('          <FeedbackName />\n')
        output_file.write('          <Modifier>0</Modifier>\n')
        output_file.write('        </Button>\n')
        output_file.write('        <Button Name="Indexed Event Log" X="3" Y="0" Width="1" Height="1">\n')
        output_file.write('          <Action>20</Action>\n')
        output_file.write('          <Link />\n')
        output_file.write('          <Endpoint />\n')
        output_file.write('          <Equipment />\n')
        output_file.write('          <ItemName />\n')
        output_file.write('          <FeedbackName />\n')
        output_file.write('          <Modifier>0</Modifier>\n')
        output_file.write('        </Button>\n')
        output_file.write('      </Menu>\n')
        output_file.write('      <Index Name="Index" X="0" Y="0" Width="5" Height="5">\n')

        build_pages(output_file, pages_list, pts_dict)

        output_file.write('      </Index>\n')
        output_file.write('    </Display>\n')
        output_file.write('    <Appearances BackColor="DarkGreen" ForeColor="White">\n')
        output_file.write('      <Static BackColor="Cyan" ForeColor="Black" />\n')
        output_file.write('      <Quality BackColor="Lime" ForeColor="Black">\n')
        output_file.write('        <Offline BackColor="Blue" ForeColor="LightCyan" />\n')
        output_file.write('        <CommFail BackColor="LightCyan" ForeColor="Blue" />\n')
        output_file.write('        <ManuallyEntered BackColor="White" ForeColor="Black" />\n')
        output_file.write('      </Quality>\n')
        output_file.write('      <Alarm BackColor="Red" ForeColor="Black" />\n')
        output_file.write('    </Appearances>\n')
        output_file.write('  </DisplaySection>\n')
        output_file.write('</AlarmConfiguration>')
