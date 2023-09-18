from pymongo_get_database import get_database
dbname = get_database('calst_test')
collection_name = dbname["Spanish"]
query = {"$and": [{"Lesson":"Lesson 2"},{"Group lesson": {"$exists": True}}]}
for doc1 in collection_name.find(query):
   print(doc1['ExerciseName'],doc1['Group lesson'] )

for i in collection_name.distinct('Lesson',{'Level':'beginner'}):
    print(i)
    print(collection_name.distinct('Group lesson', {'Lesson': str(i)}))

query = {'MP_wordpairs': { '$elemMatch': {'1.word':'toco'}}}
for doc1 in collection_name.find(query):
   print(doc1['ExerciseName'], 'here')


query = {'MP_wordpairs.$*.word': 'toco'}
for doc1 in collection_name.find(query):
   print(doc1['ExerciseName'], 'here1')

results=collection_name.find({'ExerciseName': "p-t"})
for i in results:
   for j in i['MP_wordpairs']:
      ind = len(j)
      print(ind)
      for ex in (0,ind-1):
         print(j[ex]['word'])