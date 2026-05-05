from dataclasses import dataclass, field
import re
from openpyxl import Workbook


@dataclass
class Disciplina:
    turno: str
    codigo: str
    nome: str
    CH: int
    tipo: str
    horario: dict[str, list[str]] #a parte mais importante da estrutura disciplina, sendo a chave o dia da semana e os valores os horários, em código
    periodo: int
    media: float #não sei se devo manter esse dado
    requisito: set[str] #mudei de list[str] para set[str] por motivos de pesquisa O[1] ao usar as funções do set utilizando o calculo hach pra acessar o endereço e não internar O[n]
    co_requisito: set[str] #as únicas 2 disciplinas que cai nessa situação de co_requisito é P3 e P2, devemos levar em conta?
    

@dataclass
class Aluno:
    nome: str
    matricula: str
    periodos_feitos: int
    dcp_feita: set[str] #somente as disciplinas que estão marcadas como concluidas, na extração do pdf
    dcp_fazer: set[str] #todas as disciplinas que o aluno tem que fazer, como ele pode mudar de critérios de aconselhamento, esse set vai poder ser mutável várias vezes
    
@dataclass
class Periodo:                                               
    ch_total: int = 0
    Segunda: dict = field(default_factory=lambda: {'M1': None, 'M2': None, 'M3': None, 'M4': None, 'M5': None, 'M6': None, 'T1': None, 'T2': None, 'T3': None, 'T4': None, 'T5': None, 'T6': None, 'N1': None, 'N2': None, 'N3': None})
    Terça: dict = field(default_factory=lambda: {'M1': None, 'M2': None, 'M3': None, 'M4': None, 'M5': None, 'M6': None, 'T1': None, 'T2': None, 'T3': None, 'T4': None, 'T5': None, 'T6': None, 'N1': None, 'N2': None, 'N3': None})
    Quarta: dict = field(default_factory=lambda: {'M1': None, 'M2': None, 'M3': None, 'M4': None, 'M5': None, 'M6': None, 'T1': None, 'T2': None, 'T3': None, 'T4': None, 'T5': None, 'T6': None, 'N1': None, 'N2': None, 'N3': None})
    Quinta: dict = field(default_factory=lambda: {'M1': None, 'M2': None, 'M3': None, 'M4': None, 'M5': None, 'M6': None, 'T1': None, 'T2': None, 'T3': None, 'T4': None, 'T5': None, 'T6': None, 'N1': None, 'N2': None, 'N3': None})
    Sexta: dict = field(default_factory=lambda: {'M1': None, 'M2': None, 'M3': None, 'M4': None, 'M5': None, 'M6': None, 'T1': None, 'T2': None, 'T3': None, 'T4': None, 'T5': None, 'T6': None, 'N1': None, 'N2': None, 'N3': None})
    Sábado: dict = field(default_factory=lambda: {'M1': None, 'M2': None, 'M3': None, 'M4': None, 'M5': None, 'M6': None, 'T1': None, 'T2': None, 'T3': None, 'T4': None, 'T5': None, 'T6': None, 'N1': None, 'N2': None, 'N3': None})


aluno1 = Aluno(
    nome="Teste da silva",
    matricula="202407385",
    periodos_feitos=4,
    dcp_feita={"comp388"},
    dcp_fazer={"comp389", "comp368", "comp444", "comp377"}
)


dcp_colocar = [
    Disciplina(
        turno="T",
        codigo="comp377",
        nome="Dcp teste",
        CH=60,
        tipo="OBG",
        horario = {
        "Segunda" : ["M1", "M2"],
        "Terça" : ["M1", "M2"]
        },
        periodo = 3,
        media = 0.0,
        requisito = {"comp368"},
        co_requisito = set()
    ),
    
    Disciplina(
        turno="T",
        codigo="comp312",
        nome="Dcp teste2",
        CH=60,
        tipo="OBG",
        horario = {
        "Segunda" : ["M3", "M4"],
        "Terça" : ["M3", "M4"]
        },
        periodo = 5,
        media = 0.0,
        requisito = set(),
        co_requisito = set()
    ),

    Disciplina(
        turno="T",
        codigo="comp444",
        nome="Dcp teste3",
        CH=60,
        tipo="OBG",
        horario = {
        "Segunda" : ["M4", "M5"],
        "Terça" : ["M5", "M6"]
        },
        periodo = 5,
        media = 0.0,
        requisito = set(),
        co_requisito = set()
    )
]

#Isso será a lista das disciplinas que serão colocadas no período, o negócio é se atentar na sequência de disciplinas, no caso
# a ideia é as disciplinas com mais prioridade serão as primeiras, porém como a aplicação vai ter como o próprio aluno inserir e remover disciplinas
# que ele vai cursar, ai já bagunça a lista de prioridade, tem que ver como essa lógica vai funcionar, vai ser uma lista encadeada? como?
# a informação "periodo = 5" já trás uma deia de o quanto ela é importante, juntamernte com seu tipop se é obrigatória, se é eletiva, se é extensão [ACE]
# analisar dessa forma

periodo_aconselhar = Periodo()
#Instancia a classe periodo, que é o campo do aconselhamento


#agora iremos fazer os critéros para inserir o código de disciplina nos horários dos dicionários das aulas em periodo

#critérios: verificar antes de adicionar a disciplina se:
        #1 - A soma de periodo_aconselharo.ch_total + dcp_colocar[n].ch FOR MENOR que LIMITE DE CH *nota: o limite máximo é 540 mas esse valor vai ser informado na hora das prefêrencias, se quer fazer 5 matéria, 4, 7*
        #2 - TODOS os códigos de disciplinas em dcp_colocar[n].requisito ESTÃO CONTIDOS EM aluno1.dcp_feita
        #3 - TODOS os valores dos dicionários em periodo_aconselhar cujo valores são os horários de dcp_colocar[n].horario ESTÃO VAZIOS
  


#esses são os 3 critérios mínimos para permiti a adição de uma disciplina

#com as 2 disciplinas de exemplo, o resultado esperado é:
    # para a comp377: NÃO irá ser adicionada porque o requisito comp368 não está contido em aluno1.dcp_feita
    # para a comp312: IRÁ ser adicionada porque não tem requisito e os horários estão vazios
    # para a comp444: NÃO irá ser adicionada porque houve choque de horário com a comp312, que já foi adicionada [por isso a importância da sequência de prioridade]

ch_limite = 540


for dcp in dcp_colocar:
   #critério 1
   if periodo_aconselhar.ch_total + dcp.CH <= ch_limite:   
   #é menor ou igual ao limite de CH, entra.

        #critéro 2
        if dcp.requisito.issubset(aluno1.dcp_feita):
        #o issubset verifica se todos os codigos de requisito estão contidos em aluno1.dcp_feita [se é true]

            #critério 3
            n_pode_colocar_horario = 0

            for dia, horarios_codigos in dcp.horario.items():
            #pega o dicionário do dia (segunda, terça, etc) + os horarios em código (M1, M2...), SOMENTE os dias e códigos da disciplina

                agenda_dia = getattr(periodo_aconselhar, dia)
                #pega o nome do dia (dicionário) e transforma em string

                for horario in horarios_codigos:
                    if agenda_dia[horario] is not None:
                    #verifica se o horário já tem alguma disciplina, se tiver, não pode adicionar a disciplina atual

                        n_pode_colocar_horario = 1
                        break
                        #saí do loop dos horários

                    else:
                        continue
                        #nesse momento o if é apenas a condição de verificar se algum horário não está disponível.
                        #quando acaba todos os horarios in horarios_codigo, ou encontra um break, ele sai do for e entra na linha 150 e verifica se precisa sair do for de dias [os dicionários - segunda etc] ou continuar

                if n_pode_colocar_horario == 1:
                    break
                    #saí do loop dos dias, porque já tem um horário que não pode colocar a disciplina, então precisa ir pra proxima disciplina

            if n_pode_colocar_horario == 0:
            #significa que passou por todos os dias e horários e não encontrou nehum choque de horário, agora iremos repetir o código para colocar os códigos de matéria dentro dos dicionários.
                for dia, horarios_codigos in dcp.horario.items():

                    agenda_dia = getattr(periodo_aconselhar, dia)

                    for horario in horarios_codigos:
                        agenda_dia[horario] = dcp.codigo
                        #coloca o código da disciplina no horário do dicionário do dia

                periodo_aconselhar.ch_total += dcp.CH
                #atualiza a CH total do período
        
print(periodo_aconselhar)           



