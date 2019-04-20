import tensorflow as tf
import numpy as np

tf.set_random_seed(1)
np.random.seed(1)

x = np.linspace(-1,1,300)[:,np.newaxis]
noise = np.random.normal(0,0.1,size=x.shape)
y = np.power(x,2)+noise

x = tf.placeholder(dtype=tf.float32,shape=[None,1])
y = tf.placeholder(dtype=tf.float32,shape=[None,1])

layer1 = tf.layers.dense(x,10,tf.nn.relu)
output = tf.layers.dense(layer1,1,None)

loss = tf.losses.mean_squared_error(output,y)
optimizer = tf.train.GradientDescentOptimizer(learning_rate=0.1)
train_op = optimizer.minimize(loss)

sess = tf.Session()
for step in range(100):
  sess.run(train_op)
