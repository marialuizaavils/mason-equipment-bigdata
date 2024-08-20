from pyspark.sql import SparkSession
from pyspark.sql.functions import col, sum as _sum
import matplotlib.pyplot as plt

# Iniciar uma sessão Spark
spark = SparkSession.builder.appName("MasonEquipment").getOrCreate()

# Carregar os dados do Excel usando pandas e convertê-los para um DataFrame Spark
import pandas as pd

file_path = 'Programa Excelencia Komatsu 2024.xlsx'
pdf = pd.read_excel(file_path, index_col=0)

df = spark.createDataFrame(pdf.reset_index())

# Adicionar uma coluna com o total de pontos positivos(OK), pendentes e em andamento
df = df.withColumn("Total de Equipamentos", 
                   col("positivos") + col("status") + col("ok") + 
                   col("pendentes") + col("status") + col("pendentes") + 
                   col("em andamento") + col("status"))

# Coletar os dados manipulados de volta para um DataFrame pandas para visualização
pdf_result = df.toPandas()

# Gerar o gráfico
plt.figure(figsize=(10, 6))
plt.bar(pdf_result['index'], pdf_result['Total de pontos OK'], color='skyred')
plt.xlabel('Status')
plt.ylabel('Total de pontos pendentes')
plt.title('Total de pontos em andamento')
plt.xticks(rotation=59)
plt.tight_layout()

# Salvar o gráfico
plt.savefig('pontos_programa_mason.png')

# Mostrar o gráfico
plt.show()

# Parar a sessão Spark
spark.stop()
