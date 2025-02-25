import requests

url = "https://geo-location-api-india.p.rapidapi.com/advancedmaps/v1/sy6sdjdfa2xozaqtdc6wgmdscc13j5ob/rev_geocode"

headers = {
	"x-rapidapi-key": "225b5c8c25msh7fbcb3513b2f9f8p1632b4jsncd2baa1436b1",
	"x-rapidapi-host": "geo-location-api-india.p.rapidapi.com"
}

response = requests.get(url, headers=headers)

print(response.json())