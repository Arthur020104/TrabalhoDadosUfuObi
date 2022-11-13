import pandas as pd
import urllib.request
import os
def main():
  dados = []
  for ano in range(1999,2022):
    if ano in [2003,2004]:
      continue
    if ano < 2003:
      html = urllib.request.urlopen(f"https://olimpiada.ic.unicamp.br/passadas/OBI{ano}/qmerito/p0/")
    elif ano > 2013 and ano < 2021:
      html = urllib.request.urlopen(f"https://olimpiada.ic.unicamp.br/passadas/OBI{ano}/qmerito/pu/")
    elif ano == 2021:
       html = urllib.request.urlopen(f"https://olimpiada.ic.unicamp.br/passadas/OBI{ano}/qmerito/ps/")
    else:
      html = urllib.request.urlopen(f"https://olimpiada.ic.unicamp.br/passadas/OBI{ano}/qmerito/p2/")
    html = html.read().decode()
    num = html.find('<tr class="row-header">')
    index_st_table = html.find('<table class="simple">',(num-26))
    index_end_table = html.find('</table>',index_st_table)+8
    if ano < 2003 or (ano > 2011 and ano < 2018):
      index_st_table = html.find('<table ')
      index_end_table = html.find('</table>',index_st_table)+8
    html = html[index_st_table:index_end_table]
    try:
      dado = pd.read_html(html)
      if ano < 2005:
        dado[0]["NaN"] = "NF"
        dado[0]["Pontos"] = 0
        dado[0] = dado[0].rename(columns={"Competidor(a)":"Nome"})
        #print(dado[0])
      elif ano in [2005, 2006]:
        dado[0] = dado[0].rename(columns={0:'NaN',1:"Classif.",2:"Nome",3:"Escola",4:"Cidade",5:"Estado"})
        dado[0]["Pontos"] = 0
        dado[0] = dado[0].drop(index=[0])
        print(dado[0])
      elif ano >= 2013 and ano <= 2017:
        dado[0] = dado[0].rename(columns={"Classif.": 'NaN',"Classif..1":"Classif.","Nota":"Pontos"})
      else:
        dado[0] = dado[0].rename(columns={0:'NaN',1:"Classif.",2:"Pontos",3:"Nome",4:"Escola",5:"Cidade",6:"Estado"})
        dado[0] = dado[0].drop(index=[0])
      dado[0]["Ano"] = ano
      dados.append(dado[0])
    except:
      print(ano)
  arquivo = open(f"C:/Users/{os.getlogin()}/Desktop/OBIdados.xlsx", 'w+')
  arquivo.close()
  df = pd.concat(dados).fillna(0)
  df = df.sort_values(by=['Ano','Classif.','Nome'])
  print(df)
  df.to_excel(f'C:/Users/{os.getlogin()}/Desktop/OBIdados.xlsx', index=False)


if __name__ == "__main__" :
  main()
