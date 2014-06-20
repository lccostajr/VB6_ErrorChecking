'''
Created on Jun 18, 2014

# VB6_ErrorChecking tool
# Developed by Luiz Carlos Junior

'''

import os
import sys
import shutil

def print_help():
    print("Usage: VB6_ErrorChecking [OPTIONS]")
    print("Parses a Visual Basic 6 project including error checking in all routines.")
    print()
    print("  -p, --project <PATH>  path to *.vbp project file")

def get_routine_type(line):
    rtype = None
    pos   = -1
    routine_types = ['DECLARE ', 'SUB ', 'FUNCTION ']
    for routine_type in routine_types:
        pos = line.upper().find(routine_type)
        if pos >= 0:
            rtype = routine_type[:-1]
            break
    return (rtype, pos)

if __name__ == '__main__':

    #---------------------------------------------------------------------------
    # Initialization
    #---------------------------------------------------------------------------
    proj_file = ""
    proj_path = os.getcwd()

    #---------------------------------------------------------------------------
    # Parse arguments
    #---------------------------------------------------------------------------
    i = 1
    while i < len(sys.argv):
        arg = sys.argv[i]
        i += 1
        if arg in ['-p', '--project']:
            proj_file = sys.argv[i]
            i += 1

        elif arg in ['?', '-h', '--help']:
            print_help()
            exit(0)

        else:
            print(("*** Error: unknown argument: %s\n" % arg))
            print_help()
            exit(1)

    if proj_file == "":
        print_help()
        exit(1)
    else:
        (proj_path, proj_file) = os.path.split(proj_file)

    # Prints header
    print('-' * 80)
    print(" VB6_ErrorChecking tool v1.0")
    print('-' * 80)

    #---------------------------------------------------------------------------
    # Create/clean-up destination directory
    #---------------------------------------------------------------------------
    dest_path = os.path.join(proj_path,"VB6_ErrorChecking")

    print("Project file         : %s" % proj_file)
    print("Destination directory: %s" % dest_path)
    print()

    print("Preparing destination directory")
    if not os.path.exists(dest_path):
        os.makedirs(dest_path)
    else:
        for the_file in os.listdir(dest_path):
            file_path = os.path.join(dest_path, the_file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(e)

    # Copy all files to destination
    print("Copying files to destination directory")
    for the_file in os.listdir(proj_path):
        file_src = os.path.join(proj_path, the_file)
        file_dst = os.path.join(dest_path, the_file)
        try:
            if os.path.isfile(file_src):
                shutil.copy2(file_src, file_dst)
        except Exception as e:
            print(e)

    #---------------------------------------------------------------------------
    # Parse VB project file & load file list
    #---------------------------------------------------------------------------
    print("Parsing %s project file" % proj_file)
    file_list = []
    f = open(os.path.join(dest_path,proj_file), 'r')
    content = f.readlines()
    for line in content:
        if line[0:line.find("=")] in ['Module', 'Class']:
            file_list.append(line[line.find(";")+2:].strip())
        if line[0:line.find("=")] in ['Form']:
            file_list.append(line[5:].strip())
    content.insert(1,"Module=VBE_Error_Catch; VBE_Error_Catch.bas\n")
    f.close()

    # Save file
    f = open(os.path.join(dest_path,proj_file), 'w')
    for line in content:
        f.write("%s" % line)
    f.close

    # Create VBE_Error_Catch file
    content = """
Attribute VB_Name = "VBE_Error_Catch"\n
Public Sub VBE_Error_Report(ByVal VBE_routin As String, ByVal VBE_file As String, VBE_err As ErrObject)
  Dim Lu As Integer
  Lu = 422
  Open App.Path & "\\vbe.log" For Append As #Lu
  Print #Lu, "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "." & Right(Format(Timer, "#0.00"), 2) & "] " _
           & "File: " & VBE_file _
           & "; Method: " & VBE_routin _
           & "; Description: %s%s" & VBE_err.Description & "%s" _
           & "; Source: " & VBE_err.Source
  Close #Lu
End Sub
""" % ('"', '"', '"')
    f = open(os.path.join(dest_path,"VBE_Error_Catch.bas"), 'w')
    f.writelines(content)
    f.close

    #---------------------------------------------------------------------------
    # Parse file list
    #---------------------------------------------------------------------------
    k = 0
    list_len = len(file_list)
    #for file in ["ListViewHeaders.bas"]:
    for file in file_list:
        k += 1
        print(" - Converting file %i of %i: %s" % (k, list_len, file))

        # Get content
        f = open(os.path.join(dest_path,file), 'r')
        content = f.readlines()
        f.close()

        # Initialize state
        rtype = None
        rname = None
        rend  = None
        inside_routine = False

        # Parse file
        i = 0
        new_statment = True
        statment = ""
        while i < len(content):
            line = content[i].strip()
            i += 1

            if new_statment:
                statment = ""

            # Exclude comments
            pos = line.find("'") if line.find("'") > -1 else len(line)
            statment = statment + line[0:pos]

            # Check continuation line
            if len(line) > 0:
                new_statment = line[-1] != "_"
                if not new_statment:
                    statment = statment[:-1]
                    continue

            if inside_routine:
                pos = statment.upper().find(rend)
                if pos >= 0:
                    # Define VBE_Error_Catch
                    src = "VBE_Error_Catch:\n" \
                        + "  If err Then\n" \
                        + "    Call VBE_Error_Report(VBE_routin,VBE_file,Err)\n" \
                        + "  End If\n"
                    content.insert(i-1,src)

                    # Reset state
                    rtype = None
                    rname = None
                    rend  = None
                    inside_routine = False
            else:
                (rtype, pos) = get_routine_type(statment)
                if not rtype in [None, "DECLARE"]:
                    k1 = pos + statment[pos:].find(" ")
                    k2 = pos + statment[pos:].find("(")
                    rname = statment[k1:k2].strip()
                    rend  = "END " + rtype
                    inside_routine = True

                    # Define VBE_Error_Catch
                    src = """
On Error GoTo VBE_Error_Catch
Dim VBE_routin As String
Dim VBE_file   As String
VBE_file = "%s"
VBE_routin = "%s"
""" % (file,rname)
                    content.insert(i,src)
                    i += 1

        # Save file
        f = open(os.path.join(dest_path,file), 'w')
        for line in content:
            f.write("%s" % line)
        f.close

    print()
    print("Converted project created: %s" % dest_path)
    print()
    print("Done!")
