Attribute VB_Name = "ModVar"
Option Explicit
Public Card(0 To 25) As Integer, CardB(0 To 25) As Boolean '��� ������ ���� ��
Public PlayerLive(0 To 3) As Boolean '��ҳ� �׾���
Public LastPick As Integer
Public Turn As Integer '����

Public PlaCardVal(0 To 3) As Integer '�÷��̾ ���� ī�� ����
Public Placard(0 To 3, 1 To 14) As Integer '�� �÷��̾��� ��
Public i As Integer, j As Integer, k As Integer, num As Integer '�Ϲ��� ��� ����
Public CardBool(0 To 25) As Boolean 'ī�尡 �̹� �������� ����
Public TheAnswer As Integer 'Ŭ���� ����
Public PlacardB(0 To 3, 1 To 14) As Boolean '���������� ����
Public ClickCard As Integer 'Ŭ���� ī��
Public Star As Boolean '�����ߴ��� ����
Public PlaRestCard(0 To 3) As Integer '�÷��̾� ���� ī�� ����

Public RestW As Integer, RestB As Integer '���� ī�� ����
Public ForPick As Integer, a As Integer 'ī��̱�� ����
Public GoPage As Integer '0ī��̱� 1�����ϱ�
Public PickPla As Integer '���õ� �÷��̾�
Public TurnHelp As Integer '�ϳѱ�� �����
 
Public XSize As Integer, YSize As Integer '�ػ�
