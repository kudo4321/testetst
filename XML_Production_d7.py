# Required libraries
import pandas as pd
import os
import re  # regex 正規表現（RegExp）

excel_path = r'C:\Users\flapl_000\Desktop\study\XML_attribute_ver1.0.xlsx'
file_path = r'C:\Users\flapl_000\Desktop\study\Before-kai.txt'

def read_excel_to_list(excel_path):
    # Read excel file
    df = pd.read_excel(excel_path)
    
    id_list = df.iloc[0:, 0].tolist()
    max_id = max(id_list) # 最大IDを取得
    
    # 最大IDまでのデータを取得
    df = df[df.iloc[:, 0] <= max_id] 

    # Get data from the second row for each column
    t_id = df.iloc[0:, 2].tolist() # C column is at index 2
    object_id = df.iloc[0:, 3].tolist() # D column is at index 3
    name = df.iloc[0:, 4].tolist() # E column is at index 4
    dia_1 = df.iloc[0:, 5].tolist() # F column is at index 5
    dia_2 = df.iloc[0:, 6].tolist() # G column is at index 6
    length = df.iloc[0:, 7].tolist() # H column is at index 7
    X1 = df.iloc[0:, 8].tolist() # I column is at index 8
    Y1 = df.iloc[0:, 9].tolist() # J column is at index 9
    Z1 = df.iloc[0:, 10].tolist() # K column is at index 10
    X2 = df.iloc[0:, 11].tolist() # L column is at index 11
    Y2 = df.iloc[0:, 12].tolist() # M column is at index 12
    Z2 = df.iloc[0:, 13].tolist() # N column is at index 13
    bend_angle = df.iloc[0:, 14].tolist() # O column is at index 14
    bend_radius = df.iloc[0:, 15].tolist() # P column is at index 15
    
    return t_id, object_id, name, dia_1, dia_2, length, X1, Y1, Z1, X2, Y2, Z2, bend_angle, bend_radius


def modify_txt_file(file_path):
    # Read the text file
    with open(file_path, 'r') as file:
        data = file.read()
    
    # The specific tt ids
    t_id = [459, 16193, 16543, 343995, 363165, 383055, 611652, 348015, 636362, 635339]
    
    # Find the positions of <template_table> and </template_table>
    start = data.find('<template_table>') + len('<template_table>\n')
    end = data.find('</template_table>')

    # Remove duplicate ids
    t_id = list(set(t_id))

    # Insert tt id between <template_table> and </template_table>
    new_data = data[:start]
    for id in t_id:
        new_data += f'        <tt id="{id}"/>\n' # Increased the indentation for tt ids
    new_data += '    ' + data[end:] # Increased the indentation for </template_table>


    # Create a new path for the modified file
    new_path = os.path.join(os.path.dirname(file_path), os.path.basename(file_path).split('.')[0] + '_d1.txt')

    # Write the modified data to a new text file
    with open(new_path, 'w') as file:
        file.write(new_data)
        
    return new_path


def insert_template(file_path, excel_data):
    with open(file_path, 'r') as file:
        data_lines = file.readlines()
        
    insertion_index = None
    for i, line in enumerate(data_lines):
        if '</object_table>' in line:
            insertion_index = i
            break

    if insertion_index is None:
        raise Exception('</object_table> tag not found in file.')

    for t_id, object_id, name, dia_1, dia_2, length, X1, Y1, Z1, X2, Y2, Z2, bend_angle, bend_radius in zip(*excel_data):
        object_id = str(object_id)
        name = str(name)
        
        object_id1 = object_id + str(1)
        object_id2 = object_id + str(2)
        circle1 = name + "_circle1"
        circle2 = name + "_circle2"

        #if Straight_Pipe,use this template 
        if t_id == 343995:
            template=f"""        <object id="{object_id}" name="{name}" t_id="343995">
            <value>
                <va a_id="344993" on="true" text="def" unitID="0"/>
                <va a_id="345048" text="def" unitID="0"/>
                <va a_id="345083" text="def" unitID="0"/>
                <va a_id="345111" text="def" unitID="0"/>
                <va a_id="345139" text="def" unitID="0"/>
                <va a_id="345167" text="" unitID="0"/>
                <va a_id="352006" text="1.0" unitID="3"/>
                <va a_id="363376" on="true" unitID="0"><![CDATA[{circle1}|{circle2}]]></va>
                <va a_id="389537" text="{X1}" unitID="3"/>
                <va a_id="389540" text="{Y1}" unitID="3"/>
                <va a_id="389543" text="{Z1}" unitID="3"/>
                <va a_id="389566" text="{X2}" unitID="0"/>
                <va a_id="389569" text="{Y2}" unitID="0"/>
                <va a_id="389572" text="{Z2}" unitID="0"/>
                <va a_id="389595" text="0" unitID="0"/>
                <va a_id="389599" text="0" unitID="0"/>
                <va a_id="389603" text="0" unitID="0"/>
                <va a_id="390620" text="h578792" unitID="0"/>
                <va a_id="390623" text="0" unitID="0"/>
                <va a_id="436718" unitID="3"><![CDATA[{length}|0.0]]></va>
                <va a_id="437079" unitID="3"><![CDATA[0.0|0.0]]></va>
                <va a_id="437083" unitID="3"><![CDATA[0.0|0.0]]></va>
                <va a_id="545286" text="" unitID="3"/>
                <va a_id="545294" text="" unitID="32"/>
                <va a_id="545349" text="ign" unitID="0"/>
                <va a_id="545376" text="" unitID="0"/>
                <va a_id="545382" text="h545413" unitID="0"/>
                <va a_id="579827" text="h579829" unitID="0"/>
                <va a_id="888437" text="ign" unitID="0"/>
                <va a_id="888438" text="" unitID="53"/>
                <va a_id="890031" text="ign" unitID="0"/>
                <va a_id="890049" text="ign" unitID="102"/>
                <va a_id="951205" text="def" unitID="3"/>
                <va a_id="951706" text="" unitID="3"/>
                <va a_id="966264" unitID="3"/>
                <va a_id="968288" text="" unitID="0"/>
                <va a_id="968289" text="" unitID="0"/>
                <va a_id="968290" text="" unitID="0"/>
                <va a_id="968291" text="" unitID="0"/>
                <va a_id="968292" text="" unitID="0"/>
                <va a_id="968293" text="" unitID="0"/>
                <va a_id="968294" on="true" text="" unitID="32"/>
                <va a_id="968295" text="" unitID="0"/>
                <va a_id="968296" text="" unitID="0"/>
                <va a_id="968297" text="" unitID="0"/>
                <va a_id="968299" text="" unitID="0"/>
                <va a_id="968300" on="true" text="" unitID="0"/>
                <va a_id="968301" text="" unitID="0"/>
                <va a_id="971891" on="true" text="h971632" unitID="0"/>
                <va a_id="971892" text="" unitID="161"/>
                <va a_id="972481" on="true" text="" unitID="0"/>
                <va a_id="972730" text="" unitID="0"/>
                <va a_id="976602" text="" unitID="0"/>
                <va a_id="976603" text="" unitID="0"/>
                <va a_id="977636" text="def" unitID="25"/>
                <va a_id="977638" text="" unitID="29"/>
                <va a_id="977639" text="" unitID="0"/>
                <va a_id="985282" text="" unitID="0"/>
                <va a_id="985283" on="true" text="" unitID="0"/>
                <va a_id="985286" text="" unitID="0"/>
                <va a_id="985287" text="" unitID="3"/>
                <va a_id="985288" text="" unitID="61"/>
                <va a_id="985289" text="" unitID="61"/>
                <va a_id="985290" text="" unitID="61"/>
                <va a_id="985291" text="" unitID="0"/>
                <va a_id="985292" text="" unitID="3"/>
                <va a_id="985293" text="" unitID="3"/>
                <va a_id="985294" text="" unitID="3"/>
                <va a_id="985295" text="" unitID="3"/>
                <va a_id="985296" text="" unitID="0"/>
                <va a_id="985297" text="def" unitID="3"/>
                <va a_id="985298" text="def" unitID="3"/>
                <va a_id="985299" text="def" unitID="3"/>
                <va a_id="985300" text="def" unitID="3"/>
                <va a_id="985301" text="def" unitID="3"/>
                <va a_id="985302" text="def" unitID="3"/>
                <va a_id="985303" text="" unitID="0"/>
                <va a_id="985304" on="true" text="" unitID="0"/>
                <va a_id="985305" text="h971124" unitID="0"/>
                <va a_id="985306" text="ign" unitID="0"/>
                <va a_id="985307" text="" unitID="0"/>
                <va a_id="985308" text="" unitID="0"/>
                <va a_id="985309" text="" unitID="0"/>
                <va a_id="985310" text="" unitID="61"/>
                <va a_id="985311" text="" unitID="139"/>
                <va a_id="985312" text="" unitID="58"/>
                <va a_id="985313" text="" unitID="53"/>
                <va a_id="985314" text="" unitID="53"/>
                <va a_id="985315" text="ign" unitID="61"/>
                <va a_id="985316" text="" unitID="0"/>
                <va a_id="985317" text="" unitID="0"/>
                <va a_id="985318" text="BodyMotion" unitID="0"/>
                <va a_id="1004876" text="" unitID="3"/>
                <va a_id="1004877" text="" unitID="3"/>
                <va a_id="1004878" text="h978517" unitID="0"/>
                <va a_id="1004879" text="def" unitID="0"/>
                <va a_id="1004880" text="def" unitID="0"/>
                <va a_id="1004881" text="def" unitID="0"/>
                <va a_id="1008251" text="" unitID="0"/>
            </value>
        </object>
        <object id="{object_id1}" name="{circle1}" t_id="363165">
            <value>
                <va a_id="457628" text="{dia_1}" unitID="3"/>
                <va a_id="490449" text="0" unitID="2"/>
                <va a_id="490452" text="0" unitID="2"/>
            </value>
        </object>
        <object id="{object_id2}" name="{circle2}" t_id="363165">
            <value>
                <va a_id="457628" text="{dia_2}" unitID="3"/>
                <va a_id="490449" text="0" unitID="2"/>
                <va a_id="490452" text="0" unitID="2"/>
            </value>
        </object>
"""

        #if Bend_Pipe,use this template
        elif t_id == 348015:
            template = f"""        <object id="{object_id}" name="{name}" t_id="348015">
            <value>
                <va a_id="349039" on="true" text="def" unitID="0"/>
                <va a_id="349074" text="def" unitID="0"/>
                <va a_id="349129" text="def" unitID="0"/>
                <va a_id="349157" text="def" unitID="0"/>
                <va a_id="349185" text="def" unitID="0"/>
                <va a_id="349193" text="" unitID="0"/>
                <va a_id="352071" text="1.0" unitID="3"/>
                <va a_id="352095" text="{X1}" unitID="3"/>
                <va a_id="352098" text="{Y1}" unitID="3"/>
                <va a_id="352101" text="{Z1}" unitID="3"/>
                <va a_id="352104" text="{X2}" unitID="0"/>
                <va a_id="352127" text="{Y2}" unitID="0"/>
                <va a_id="352130" text="{Z2}" unitID="0"/>
                <va a_id="352158" text="{bend_angle}" unitID="3"/>
                <va a_id="352161" text="{bend_radius}" unitID="61"/>
                <va a_id="363654" on="true" unitID="0"><![CDATA[{circle1}|{circle2}]]></va>
                <va a_id="389115" text="-0.15049295094043852" unitID="0"/>
                <va a_id="389120" text="0.8211117160711776" unitID="0"/>
                <va a_id="389122" text="0.5505700153109799" unitID="0"/>
                <va a_id="390407" text="h578637" unitID="0"/>
                <va a_id="390410" text="0" unitID="0"/>
                <va a_id="436721" unitID="3"/>
                <va a_id="436958" unitID="3"><![CDATA[0.0]]></va>
                <va a_id="436962" unitID="3"><![CDATA[0.0]]></va>
                <va a_id="545742" text="" unitID="3"/>
                <va a_id="545770" text="" unitID="32"/>
                <va a_id="545805" text="ign" unitID="0"/>
                <va a_id="545832" text="" unitID="0"/>
                <va a_id="545858" text="h545889" unitID="0"/>
                <va a_id="579702" text="h579704" unitID="0"/>
                <va a_id="825976" text="def" unitID="3"/>
                <va a_id="888441" text="ign" unitID="0"/>
                <va a_id="888442" text="" unitID="53"/>
                <va a_id="890022" text="ign" unitID="0"/>
                <va a_id="890039" text="ign" unitID="102"/>
                <va a_id="951203" text="def" unitID="3"/>
                <va a_id="951711" text="" unitID="3"/>
                <va a_id="952270" text="h950877" unitID="0"/>
                <va a_id="967147" unitID="3"/>
                <va a_id="968349" text="" unitID="0"/>
                <va a_id="968350" text="" unitID="0"/>
                <va a_id="968351" text="" unitID="0"/>
                <va a_id="968352" text="" unitID="0"/>
                <va a_id="968353" text="" unitID="0"/>
                <va a_id="968354" text="" unitID="0"/>
                <va a_id="968355" on="true" text="" unitID="32"/>
                <va a_id="968356" text="" unitID="0"/>
                <va a_id="968357" text="" unitID="0"/>
                <va a_id="968358" text="" unitID="0"/>
                <va a_id="968359" text="" unitID="0"/>
                <va a_id="968360" on="true" text="" unitID="0"/>
                <va a_id="968361" text="" unitID="0"/>
                <va a_id="971859" on="true" text="h971636" unitID="0"/>
                <va a_id="971860" text="" unitID="161"/>
                <va a_id="972482" on="true" text="" unitID="0"/>
                <va a_id="972702" text="" unitID="0"/>
                <va a_id="976568" text="" unitID="0"/>
                <va a_id="976569" text="" unitID="0"/>
                <va a_id="977640" text="def" unitID="25"/>
                <va a_id="977641" text="" unitID="29"/>
                <va a_id="977642" text="" unitID="0"/>
                <va a_id="985319" text="" unitID="0"/>
                <va a_id="985320" on="true" text="" unitID="0"/>
                <va a_id="985323" text="" unitID="0"/>
                <va a_id="985324" text="" unitID="3"/>
                <va a_id="985325" text="" unitID="61"/>
                <va a_id="985326" text="" unitID="61"/>
                <va a_id="985327" text="" unitID="61"/>
                <va a_id="985328" text="" unitID="0"/>
                <va a_id="985329" text="" unitID="3"/>
                <va a_id="985330" text="" unitID="3"/>
                <va a_id="985331" text="" unitID="3"/>
                <va a_id="985332" text="" unitID="3"/>
                <va a_id="985333" text="" unitID="0"/>
                <va a_id="985334" text="def" unitID="3"/>
                <va a_id="985335" text="def" unitID="3"/>
                <va a_id="985336" text="def" unitID="3"/>
                <va a_id="985337" text="def" unitID="3"/>
                <va a_id="985338" text="def" unitID="3"/>
                <va a_id="985339" text="def" unitID="3"/>
                <va a_id="985340" text="" unitID="0"/>
                <va a_id="985341" on="true" text="" unitID="0"/>
                <va a_id="985342" text="h971127" unitID="0"/>
                <va a_id="985343" text="ign" unitID="0"/>
                <va a_id="985344" text="" unitID="0"/>
                <va a_id="985345" text="" unitID="0"/>
                <va a_id="985346" text="" unitID="0"/>
                <va a_id="985347" text="" unitID="61"/>
                <va a_id="985348" text="" unitID="139"/>
                <va a_id="985349" text="" unitID="58"/>
                <va a_id="985350" text="" unitID="53"/>
                <va a_id="985351" text="" unitID="53"/>
                <va a_id="985352" text="ign" unitID="61"/>
                <va a_id="985353" text="" unitID="0"/>
                <va a_id="985354" text="" unitID="0"/>
                <va a_id="985355" text="BodyMotion" unitID="0"/>
                <va a_id="1004839" text="" unitID="3"/>
                <va a_id="1004840" text="" unitID="3"/>
                <va a_id="1004843" text="def" unitID="0"/>
                <va a_id="1004844" text="def" unitID="0"/>
                <va a_id="1004845" text="def" unitID="0"/>
                <va a_id="1004846" text="h978506" unitID="0"/>
                <va a_id="1008249" text="" unitID="0"/>
            </value>
        </object>
        <object id="{object_id1}" name="{circle1}" t_id="363165">
            <value>
                <va a_id="457628" text="{dia_1}" unitID="3"/>
                <va a_id="490449" text="0" unitID="2"/>
                <va a_id="490452" text="0" unitID="2"/>
            </value>
        </object>
        <object id="{object_id2}" name="{circle2}" t_id="363165">
            <value>
                <va a_id="457628" text="{dia_2}" unitID="3"/>
                <va a_id="490449" text="0" unitID="2"/>
                <va a_id="490452" text="0" unitID="2"/>
            </value>
        </object>
"""

        else:
            print(f"Unknown t_id: {t_id}. Error")
            template ="Error"
            continue

        filled_template = template.format(
                object_id=object_id,
                name=name,
                t_id =t_id,
                X1=X1,
                Y1=Y1,
                Z1=Z1,
                X2=X2,
                Y2=Y2,
                Z2=Z2,
                circle1=circle1,
                circle2=circle2,
                object_id1=object_id1,
                object_id2=object_id2,
                dia_1=dia_1,
                dia_2=dia_2,
                length=length
            )

        data_lines.insert(insertion_index, filled_template)
        insertion_index += 1  # Move the insertion index after each insertion
        
    with open(file_path, 'w') as file:
        file.write("".join(data_lines))


def insert_mesh_relation_template(file_path, excel_data):
    with open(file_path, 'r') as file:
        data = file.read()

    # Replace <comp_mesh/> with <comp_mesh></comp_mesh>
    data = data.replace('<comp_mesh/>', '<comp_mesh>\n        </comp_mesh>')

    # Save the file after replacing <comp_mesh/>
    with open(file_path, 'w') as file:
        file.write(data)

    with open(file_path, 'r') as file:
        data_lines = file.readlines()

    insertion_index = None
    for i, line in enumerate(data_lines):
        if '<comp_mesh>' in line:
            insertion_index = i + 1
            break

    if insertion_index is None:
        raise Exception('<comp_mesh> tag not found in file.')

    
    for _, object_id, name, _, _, _, _, _, _, _, _, _, _, _ in zip(*excel_data):
        object_id = str(object_id)[1:] # Remove the first character
        name = str(name)

        # Find the line with the matching name
        matching_line = None
        for line in data_lines:
            if name in line:
                matching_line = line
                break

        # Extract mesh_id from the matching line
        if matching_line:
            mesh_id_match = re.search(r'object id="([^"]+)"', matching_line)
            if mesh_id_match:
                mesh_id = mesh_id_match.group(1)[1:]
            else:
                print(f'No object id found in line: {matching_line}')
                continue
        else:
            print(f'Name "{name}" not found in file.')
            continue

        # Construct the compwmesh_template
        compwmesh_template = f"""            <compwmesh id="{object_id}">
                <mesh id="{mesh_id}"/>
            </compwmesh>
"""

        # Insert the compwmesh_template between the <comp_mesh> and </comp_mesh> tags
        data_lines.insert(insertion_index, compwmesh_template)
        insertion_index += 1  # Move the insertion index after each insertion

    # Write the modified data back to the file
    with open(file_path, 'w') as file:
        file.write("".join(data_lines))


# Activate Functions
excel_data = read_excel_to_list(excel_path)
new_path = modify_txt_file(file_path)

insert_template(new_path, excel_data)
insert_mesh_relation_template(new_path, excel_data)


# Verify the data
t_id, object_id, name, dia_1, dia_2, length, X1, Y1, Z1, X2, Y2, Z2, bend_angle, bend_radius = read_excel_to_list(excel_path)        

variables = [t_id, object_id, name, dia_1, dia_2, length, X1, Y1, Z1, X2, Y2, Z2, bend_angle, bend_radius]
variable_names = ['t_id', 'object_id', 'name', 'dia_1', 'dia_2', 'length', 'X1', 'Y1', 'Z1', 'X2', 'Y2', 'Z2', 'bend_angle', 'bend_radius']

for var_name, var in zip(variable_names, variables):
    print(f"{var_name}:")
    for val in var:
        print(val)
    print('\n')
