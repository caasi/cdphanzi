Attribute VB_Name = "�����ܼ�"
Option Explicit
Public Const mdi�~�r�r�ΥN�X = 0
Public Const Big5�r�ڥN�X = 1, Big5��²�Ʀr�r�ڥN�X = 2, �r�ڥN�X = 3, �p�f�W��r�N�X = 4, ����r�ڥN�X = 5, �Ұ���r�ڥN�X = 6, ���t²����r�r�ڥN�X = 7, ����~�r�N�X = 8
Public Const �d�������N�X = 9, ���峡���N�X = 10, ²�|�N�X = 11, �K���N�X = 12, �ϧΤ�r�N�X = 13, �c�r�Ÿ��N�X = 14
Public Const �r�δF�ťN�X = 15, �X�B�˦r�N�X = 16, �r�ε��c�N�X = 17, ����r��N�X = 18, ����r�ڥN�X = 19
Public Const �r�κt�ܥN�X = 20, �r�ί��ޥN�X = 21
Public Const �r�δF��_�˦r��� = 1, �r�δF��_�𪬵��c = 2
Public Const �X�B�˦r_�˦r��� = 1, �X�B�˦r_�𪬵��c = 2
Public Const mdi�~�r�r��_�s����� = 1, mdi�~�r�r��_�~�r����� = 2, mdi�~�r�r��_�r�Τ�� = 3
Public Const mdi�~�r�r��_�`���e��� = 4, mdi�~�r�r��_������� = 5, mdi�~�r�r��_�����������e��� = 6
Public Const mdi�~�r�r��_�`����� = 7, mdi�~�r�r��_���X��� = 8, mdi�~�r�r��_�ܾe�X��� = 9
Public Const mdi�~�r�r��_�c�r����� = 10, mdi�~�r�r��_�U�Ƥ�� = 11, mdi�~�r�r��_�զr�r�Ƥ�� = 12, mdi�~�r�r��_�զr�r�Ƨt���g��� = 13
Public Const mdi�~�r�r��_������ = 14, mdi�~�r�r��_�j�~�r��� = 15
Public Const �r�θ`�I�аO = 0, �c�r���`�I�аO = 1, ���W�`�I�аO = 2, ��L�`�I�аO = 3
Public Const �r��Ӽ� = 27
Public �ƻs����X As Boolean, �ƻsBig5�r�� As Boolean, �ƻsunicode�r�� As Boolean
Public Const ccDefault = 0, ccHourglass = 11
Public Const ²���s���`�� = 4, �w�]�s���`�� = 4
Public �t�θ�Ʈw As Database, �p�f��Ʈw As Database, �����Ʈw As Database, �Ұ����Ʈw As Database, ���t��r��Ʈw As Database
Public �`�βŸ��γ��� As Recordset, �d������ As Recordset, ���峡�� As Recordset, �`�βŸ��γ������� As Recordset
Public �r�Ϊ� As Recordset
Public �˦r�� As Recordset, ����r�� As Recordset, ���g�r�� As Recordset
Public �����˦r�� As Recordset, ���Ѳ���r�� As Recordset, ���Ѳ��g�r�� As Recordset, ���Ѧr�� As Recordset
Public �p�f�˦r�� As Recordset, �p�f����r�� As Recordset, �p�f���g�r�� As Recordset, �p�f�W��r As Recordset
Public �����˦r�� As Recordset, ����ɿ�� As Recordset, ���岧��r�� As Recordset, ���岧�g�r�� As Recordset, ���嶰�����W As Recordset, ���嶰���ޱo As Recordset, ������L As Recordset, ���岧�g�r�� As Recordset, ����r�� As Recordset
Public �Ұ����˦r�� As Recordset, �Ұ��岧��r�� As Recordset, �Ұ��岧�g�r�� As Recordset, �Ұ��岧�g�r�� As Recordset, �Ұ���r�� As Recordset
Public ���t��r�˦r�� As Recordset, ���t��r�ɿ�� As Recordset, ���t��r����r�� As Recordset, ���t��r���g�r�� As Recordset, ���t��r���g�r�� As Recordset, ���t��r�r�� As Recordset
Public �{�ε��� As String, �{�ε����N�X As Integer, �{�α���N�X As Integer
Public �{�Φr�� As String, �t�Φr�� As String, ��ܦr�� As String, ��ܦr���j�p As Integer
Public �Ϥ��ѪR�� As Long, �Ϥ��r���j�p As Long
Public �ƻs�Ϥ���Word As Boolean, �ƻs��Word���Ϥ��j�p As Integer
Public ���e�����d�� As Boolean
Public �զr�Ÿ��}�C(1 To 14) As String
Public �Ұʦr�δF�� As Boolean, �ҰʥX�B�˦r As Boolean, �Ұʦr�ε��c As Boolean, �Ұʲ���r�� As Boolean, �Ұʲ���r�� As Boolean, �Ұʦr�κt�� As Boolean, �Ұʦr�ί��� As Boolean, �Ұʳ���d�� As Boolean
Public �@�ε���(mdi�~�r�r�ΥN�X To �r�ί��ޥN�X) As String, �@�ε����N�X As Integer
Public ��ڧǼ� As Integer
Public �r��}�C(0 To �r��Ӽ�) As Variant
Public ��e As Integer
Public ���A�C�r�� As Long
Public �t�νs�� As Long
Public �����N�X(0 To 2) As Boolean
Public �r��_�`�Φr As String * 250, �r��_���j�X As String * 250, �r��_²�Ʀr As String * 250, �r��_�~�y�j�r�� As String * 250, �r��_����Ѧr As String * 250, �r��_����s As String * 250, �r��_����s�ϧΤ�r As String * 250, �r��_�Ұ���ġ As String * 250, �r��_���t²����r�s As String * 250, �r��_����r As String * 250
Public �F��open As String * 250, �F��winstate As String * 250, �F��left As String * 250, �F��top As String * 250, �F��width As String * 250, �F��height As String * 250
Public �X�Bopen As String * 250, �X�Bwinstate As String * 250, �X�Bleft As String * 250, �X�Btop As String * 250, �X�Bwidth As String * 250, �X�Bheight As String * 250
Public ���copen As String * 250, ���cwinstate As String * 250, ���cleft As String * 250, ���ctop As String * 250, ���cwidth As String * 250, ���cheight As String * 250
Public ����open As String * 250, ����winstate As String * 250, ����left As String * 250, ����top As String * 250, ����width As String * 250, ����height As String * 250
Public ����open As String * 250, ����winstate As String * 250, ����left As String * 250, ����top As String * 250, ����width As String * 250, ����height As String * 250
Public ����open As String * 250, ����winstate As String * 250, ����left As String * 250, ����top As String * 250, ����width As String * 250, ����height As String * 250
Public �t��open As String * 250, �t��winstate As String * 250, �t��left As String * 250, �t��top As String * 250, �t��width As String * 250, �t��height As String * 250
Public ����open As String * 250, ����winstate As String * 250, ����left As String * 250, ����top As String * 250, ����width As String * 250, ����height As String * 250
Public ��lfirst As String * 250, ��l���󶶧� As String * 250, ��l���g���� As String * 250, ��l�v�ŦC�X As String * 250, ��l�ѧΦC�X As String * 250, ��lcopy As String
Public ��l���F�~�y�j�r�� As String * 250, ��l�ا��~�y�j�r�� As String * 250, ��l����j��� As String * 250
Public ��l����Ѧr���L As String * 250, ��l���ػ���Ѧr As String * 250
Public ��l����s As String * 250, ��l������L As String * 250, ��l���徹�� As String * 250, ��l����ޱo As String * 250
Public ��l�Ұ�������ġ As String * 250, ��l�Ұ���r���L As String * 250, ��l�Ұ���r���� As String * 250
Public ��l���t²����r�s As String * 250, ��l���t��r�X�B As String * 250
Public ��lUnicode As String * 250, ��lBig5 As String * 250
Public ��l�Ұ���t�� As String * 250, ��l����t�� As String * 250, ��l���t��r�t�� As String * 250, ��l�p�f�t�� As String * 250
Public ��l�r�W As String * 250, ��l����X As String * 250
Public ��lCopyToWord As String * 250, ��lCopyUnicode As String * 250
Public �X�B���Ұ��� As Boolean, �X�B������ As Boolean, �X�B���p�f As Boolean, �X�B������r As Boolean
Public �X�B���Ұ���X�� As Boolean, �X�B�����ǰt As Boolean
Public �w���J�e�� As Integer
Public �w�]�s���Ҧ� As Integer, ²���s���Ҧ� As Boolean, ���ܹw�]�s�� As Boolean
Public �F�Ū��A�C As String, ���c���A�C As String, ���󪬺A�C As String, ���骬�A�C As String
Public ���O As String
Public �즲�r�� As String
Public �Ȧs�ؿ� As String, bmpcount As Long, �Ȧs���� As String, ���N��r As String

Public WordApp As Word.Application, WordWasNotRunning As Boolean
