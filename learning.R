library(tidyverse)
data(mtcars)
map(mtcars, mean) %>% head # Returns a list
map_dbl(mtcars, mean) %>% head # Returns a vector - of class double

mtcars %>% 
  split(.$cyl)

models <- mtcars %>% 
  split(.$cyl) %>% 
  map(function(df) lm(mpg ~ wt, data = df))

models

models %>% 
  map(summary) %>% 
  map_dbl(~.$r.squared)
