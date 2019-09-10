# coding: utf-8
from django.db import models


class Person(models.Model):
    name = models.CharField(max_length=100)
    age = models.IntegerField()

    def __unicode__(self):
        return self.name

    def __str__(self):
        return self.name


class Recipe(models.Model):
    name = models.CharField('Recipe name', max_length=100)
    owner = models.ForeignKey('Person', blank=True, null=True, on_delete=models.SET_NULL)
    ingredient = models.ManyToManyField('Ingredient', blank=True, through='Association', related_name='recipe')
    cooking_time = models.DurationField(blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)

    VEGAN = 'V'
    GLUTENFREE = 'G'
    RECIPETYPES = (
        (VEGAN, 'Vegan'),
        (GLUTENFREE, 'Gluten free'),
    )

    rtype = models.CharField('Type', max_length=1, choices=RECIPETYPES, blank=True, null=True)

    def __unicode__(self):
        return self.name

    def __str__(self):
        return self.name


class Ingredient(models.Model):
    name = models.CharField('Ingredient name', max_length=100)

    def __unicode__(self):
        return self.name

    def __str__(self):
        return self.name


class Association(models.Model):
    recipe = models.ForeignKey('Recipe', on_delete=models.CASCADE)
    ingredient = models.ForeignKey('Ingredient', on_delete=models.CASCADE)
    amount = models.PositiveIntegerField(null=True, blank=True)

    def __unicode__(self):
        return self.name

    def __str__(self):
        return self.name