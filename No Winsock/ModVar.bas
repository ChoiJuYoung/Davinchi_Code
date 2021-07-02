Attribute VB_Name = "ModVar"
Option Explicit
Public Card(0 To 25) As Integer, CardB(0 To 25) As Boolean '모든 숫자의 고유 값
Public PlayerLive(0 To 3) As Boolean '살았나 죽었나
Public LastPick As Integer
Public Turn As Integer '차례

Public PlaCardVal(0 To 3) As Integer '플레이어별 가진 카드 갯수
Public Placard(0 To 3, 1 To 14) As Integer '각 플레이어의 판
Public i As Integer, j As Integer, k As Integer, num As Integer '일반적 사용 변수
Public CardBool(0 To 25) As Boolean '카드가 이미 뽑혔는지 여부
Public TheAnswer As Integer '클릭시 정답
Public PlacardB(0 To 3, 1 To 14) As Boolean '맞춰졌는지 여부
Public ClickCard As Integer '클릭한 카드
Public Star As Boolean '시작했는지 여부
Public PlaRestCard(0 To 3) As Integer '플레이어 남은 카드 개수

Public RestW As Integer, RestB As Integer '남은 카드 갯수
Public ForPick As Integer, a As Integer '카드뽑기용 변수
Public GoPage As Integer '0카드뽑기 1선택하기
Public PickPla As Integer '선택된 플레이어
Public TurnHelp As Integer '턴넘기기 도움수
 
Public XSize As Integer, YSize As Integer '해상도
