

            

Statistical
Distributions

By

Bill Young




 

 


 Contents
 Tests. 4
 Shapiro-Wilk Normality
 Test 4
 Kolmogorov-Smirnov. 4
 General Information.. 4
 Discrete Distributions. 4
 Definition: 4
 With Finite Support 4
 With Infinite Support 5
 Continuous Distributions. 6
 Definition: 6
 Supported On A Bounded
 Interval 6
 Supported On Intervals Of
 Length 2π – Directional Distributions. 12
 Supported On Semi-infinite Intervals, Usually [0,∞) 17
 Supported On The Whole Real Line. 18
 With Variable Support 20
 Mixed Discrete/Continuous Distributions. 20
 Joint Distributions. 20
 Two or More Random Variables On The Sample Space. 20
 Matrix-Valued Distributions. 21
 Non-numeric Distributions. 21
 Miscellaneous Distributions. 21
 Works Cited. 22
  


 




 

 

#Tests



##Shapiro-Wilk Normality Test

 (Zaiontz, Shapiro-Wilk Original Test) (Zaiontz,
 Shapiro-Wilk Expanded Test)

##Kolmogorov-Smirnov

(Sonnier)

#General Information

Common Probability Distributions (Joyce, 2016)

 

#Discrete
Distributions



##Definition: 

“A [statistical
distribution ](http://mathworld.wolfram.com/StatisticalDistribution.html)whose
variables can take on only discrete values. Abramowitz and Stegun (1972,
p. 929) give a table of the parameters of most common discrete
distributions.” (Weisstein, Discrete Distribution, 2017)


 ##With
     Finite Support
 The [Bernoulli distribution](https://en.wikipedia.org/wiki/Bernoulli_distribution), which takes value 1 with
     probability p and value 0 with probability q = 1 − p.
  
      
      
       
       
       
       
       
       
       
       
       
       
       
       
      
      
      
     
      
      
     <img
     src="https://github.com/Temtesb/StatisticsCalculationsForExcel/blob/master/Documentation/Images/BernoulliDistribution.png"
     alt="Picture 2"> (Weisstein, Wolfram Mathworld, 2017)
 The [Rademacher distribution](https://en.wikipedia.org/wiki/Rademacher_distribution), which takes value 1 with
     probability 1/2 and value −1 with probability 1/2.
 The [binomial
     distribution](https://en.wikipedia.org/wiki/Binomial_distribution), which describes the number of
     successes in a series of independent Yes/No experiments all with the same
     probability of success.
 The [beta-binomial
     distribution](https://en.wikipedia.org/wiki/Beta-binomial_model), which describes the number of
     successes in a series of independent Yes/No experiments with heterogeneity
     in the success probability.
 The [degenerate distribution ](https://en.wikipedia.org/wiki/Degenerate_distribution)at x0, where X is certain to take
     the value x0. This does not look random, but it satisfies the definition of [random
     variable](https://en.wikipedia.org/wiki/Random_variable). This is useful because it puts
     deterministic variables and random variables in the same formalism.
 The [discrete uniform distribution](https://en.wikipedia.org/wiki/Uniform_distribution_(discrete)), where all elements of a finite [set ](https://en.wikipedia.org/wiki/Set_theory)are
     equally likely. This is the theoretical distribution model for a balanced
     coin, an unbiased die, a casino roulette, or the first card of a
     well-shuffled deck.
 The [hypergeometric distribution](https://en.wikipedia.org/wiki/Hypergeometric_distribution), which describes the number of
     successes in the first m of a series of n consecutive Yes/No
     experiments, if the total number of successes is known. This distribution
     arises when there is no replacement.
 The [Poisson binomial distribution](https://en.wikipedia.org/wiki/Poisson_binomial_distribution), which describes the number of
     successes in a series of independent Yes/No experiments with different
     success probabilities.
 [Fisher's noncentral hypergeometric
     distribution](https://en.wikipedia.org/wiki/Fisher%27s_noncentral_hypergeometric_distribution)
 [Wallenius' noncentral hypergeometric
     distribution](https://en.wikipedia.org/wiki/Wallenius%27_noncentral_hypergeometric_distribution)
 [Benford's law](https://en.wikipedia.org/wiki/Benford%27s_law), which describes the frequency of
     the first digit of many naturally occurring data.


 


 ##With
     Infinite Support
 The [beta negative binomial distribution](https://en.wikipedia.org/wiki/Beta_negative_binomial_distribution)
 The [Boltzmann distribution](https://en.wikipedia.org/wiki/Boltzmann_distribution), a discrete distribution important
     in [statistical
     physics ](https://en.wikipedia.org/wiki/Statistical_physics)which
     describes the probabilities of the various discrete energy levels of a
     system in [thermal
     equilibrium](https://en.wikipedia.org/wiki/Thermal_equilibrium). It has a continuous analogue.
     Special cases include: 
 
  The [Gibbs
      distribution](https://en.wikipedia.org/wiki/Gibbs_distribution)
  The [Maxwell–Boltzmann distribution](https://en.wikipedia.org/wiki/Maxwell%E2%80%93Boltzmann_distribution)
 
 The [Borel
     distribution](https://en.wikipedia.org/wiki/Borel_distribution)
 The [Champernowne distribution](https://en.wikipedia.org/wiki/Champernowne_distribution)
 The [extended negative binomial distribution](https://en.wikipedia.org/wiki/Extended_negative_binomial_distribution)
 The [extended hypergeometric distribution](https://en.wikipedia.org/wiki/Extended_hypergeometric_distribution)
 The [generalized log-series distribution](https://en.wikipedia.org/w/index.php?title=Generalized_log-series_distribution&action=edit&redlink=1)
 The [geometric distribution](https://en.wikipedia.org/wiki/Geometric_distribution), a discrete distribution which
     describes the number of attempts needed to get the first success in a
     series of independent Bernoulli trials, or alternatively only the number
     of losses before the first success (i.e. one less).
 The [logarithmic (series) distribution](https://en.wikipedia.org/wiki/Logarithmic_distribution)
 The [negative binomial distribution ](https://en.wikipedia.org/wiki/Negative_binomial_distribution)or
     Pascal distribution a generalization of the geometric distribution to the nth
     success.
 The discrete [compound Poisson distribution](https://en.wikipedia.org/wiki/Compound_Poisson_distribution)
 The [parabolic fractal distribution](https://en.wikipedia.org/wiki/Parabolic_fractal_distribution)
 The [Poisson
     distribution](https://en.wikipedia.org/wiki/Poisson_distribution), which describes a very large number
     of individually unlikely events that happen in a certain time interval.
     Related to this distribution are a number of other distributions: the [displaced Poisson](https://en.wikipedia.org/wiki/Displaced_Poisson_distribution), the hyper-Poisson, the general Poisson
     binomial and the Poisson type distributions. 
 
  The [Conway–Maxwell–Poisson distribution](https://en.wikipedia.org/wiki/Conway%E2%80%93Maxwell%E2%80%93Poisson_distribution), a two-parameter extension of the [Poisson
      distribution ](https://en.wikipedia.org/wiki/Poisson_distribution)with
      an adjustable rate of decay.
  The [Zero-truncated Poisson distribution](https://en.wikipedia.org/wiki/Zero-truncated_Poisson_distribution), for processes in which zero counts
      are not observed
 
 The [Polya–Eggenberger distribution](https://en.wikipedia.org/w/index.php?title=Polya%E2%80%93Eggenberger_distribution&action=edit&redlink=1)
 The [Skellam
     distribution](https://en.wikipedia.org/wiki/Skellam_distribution), the distribution of the difference
     between two independent Poisson-distributed random variables.
 The [skew elliptical distribution](https://en.wikipedia.org/w/index.php?title=Skew_elliptical_distribution&action=edit&redlink=1)
 The [Yule–Simon distribution](https://en.wikipedia.org/wiki/Yule%E2%80%93Simon_distribution)
 The [zeta
     distribution ](https://en.wikipedia.org/wiki/Zeta_distribution)has
     uses in applied statistics and statistical mechanics, and perhaps may be
     of interest to number theorists. It is the [Zipf
     distribution ](https://en.wikipedia.org/wiki/Zipf_distribution)for
     an infinite number of elements.
 [Zipf's law ](https://en.wikipedia.org/wiki/Zipf%27s_law)or
     the Zipf distribution. A discrete [power-law ](https://en.wikipedia.org/wiki/Power_law)distribution,
     the most famous example of which is the description of the frequency of
     words in the English language.
 The [Zipf–Mandelbrot
     law ](https://en.wikipedia.org/wiki/Zipf%E2%80%93Mandelbrot_law)is
     a discrete power law distribution which is a generalization of the [Zipf
     distribution](https://en.wikipedia.org/wiki/Zipf_distribution).


 

#Continuous
Distributions



##Definition:

 A continuous
random variable is a random variable with a set of possible values (known as
the range or support) that is infinite and uncountable. Probabilities of continuous random variables (X) are defined as the area under the
curve of its distribution. Thus, only ranges of
values can have a nonzero probability.  (Minitab)


 ##Supported
     On A Bounded Interval
 The [arcsine
     distribution ](https://en.wikipedia.org/wiki/Arcsine_distribution)on
     [a,b], which is a special case of the Beta distribution if a
     = 0 and b = 1.
 The [Beta
     distribution ](https://en.wikipedia.org/wiki/Beta_distribution)on
     [0,1], a family of two-parameter distributions with one mode, of which the
     uniform distribution is a special case, and which is useful in estimating
     success probabilities.
 The [logitnormal distribution ](https://en.wikipedia.org/wiki/Logitnormal)on
     (0,1).
 
      
      (Wikipedia)
 The [Dirac
     delta function ](https://en.wikipedia.org/wiki/Dirac_delta_function)although
     not strictly a function, is a limiting form of many continuous probability
     functions. It represents a discrete probability distribution
     concentrated at 0 — a [degenerate distribution ](https://en.wikipedia.org/wiki/Degenerate_distribution)— but the
     notation treats it as if it were a continuous distribution.
 The [continuous uniform distribution ](https://en.wikipedia.org/wiki/Uniform_distribution_(continuous))or
     [rectangular distribution ](https://en.wikipedia.org/wiki/Rectangular_distribution)on [a,b],
     where all points in a finite interval are equally likely.
 
      
      (Wikipedia)
 The [Irwin–Hall distribution ](https://en.wikipedia.org/wiki/Irwin%E2%80%93Hall_distribution)is
     the distribution of the sum of n i.i.d. U(0,1) random variables.
 
      
      (Wikipedia)
 The [Bates
     distribution ](https://en.wikipedia.org/wiki/Bates_distribution)is
     the distribution of the mean of n i.i.d. U(0,1) random variables.
 The [Kent
     distribution ](https://en.wikipedia.org/wiki/Kent_distribution)on
     the three-dimensional sphere.
 
      
      (Wikipedia)
 The [Kumaraswamy distribution ](https://en.wikipedia.org/wiki/Kumaraswamy_distribution)is as
     versatile as the Beta distribution but has simple closed forms for both
     the cdf and the pdf.
 
      
      (Wikipedia)
 The [logarithmic distribution (continuous)](https://en.wikipedia.org/w/index.php?title=Logarithmic_distribution_(continuous)&action=edit&redlink=1)
 The [Marchenko–Pastur distribution ](https://en.wikipedia.org/wiki/Marchenko%E2%80%93Pastur_distribution)is
     important in the theory of [random
     matrices](https://en.wikipedia.org/wiki/Random_matrices).
 The PERT distribution is a special
     case of the [beta
     distribution](https://en.wikipedia.org/wiki/Beta_distribution)
 The [raised cosine distribution](https://en.wikipedia.org/wiki/Raised_cosine_distribution)
 
      
      (Wikipedia)
 The [reciprocal distribution](https://en.wikipedia.org/wiki/Reciprocal_distribution)
 The [triangular distribution ](https://en.wikipedia.org/wiki/Triangular_distribution)on [a,
     b], a special case of which is the distribution of the sum of two
     independent uniformly distributed random variables (the convolution
     of two uniform distributions).
 
      
      (Wikipedia)
 The [trapezoidal distribution](https://en.wikipedia.org/wiki/Trapezoidal_distribution)
 The [truncated normal distribution ](https://en.wikipedia.org/wiki/Truncated_normal_distribution)on
     [a, b].
 The [U-quadratic distribution ](https://en.wikipedia.org/wiki/U-quadratic_distribution)on [a,
     b].
 The [von Mises-Fisher distribution ](https://en.wikipedia.org/wiki/Von_Mises-Fisher_distribution)on
     the N-dimensional sphere has the [von Mises distribution ](https://en.wikipedia.org/wiki/Von_Mises_distribution)as a special
     case.
 The [Wigner semicircle distribution ](https://en.wikipedia.org/wiki/Wigner_semicircle_distribution)is
     important in the theory of [random
     matrices](https://en.wikipedia.org/wiki/Random_matrices).


 


 ##Supported
     On Intervals Of Length 2π – Directional Distributions
 The [von Mises distribution](https://en.wikipedia.org/wiki/Von_Mises_distribution)
 
      
      (Wikipedia)
 The [wrapped normal distribution](https://en.wikipedia.org/wiki/Wrapped_normal_distribution)
 
      
      (Wikipedia)
 The [wrapped exponential distribution](https://en.wikipedia.org/wiki/Wrapped_exponential_distribution)
 
      
      (Wikipedia)
 The [wrapped Lévy distribution](https://en.wikipedia.org/wiki/Wrapped_L%C3%A9vy_distribution)
 The [wrapped Cauchy distribution](https://en.wikipedia.org/wiki/Wrapped_Cauchy_distribution)
 
      
      (Wikipedia)
 The [wrapped Laplace distribution](https://en.wikipedia.org/wiki/Wrapped_Laplace_distribution)
 The [wrapped asymmetric Laplace distribution](https://en.wikipedia.org/wiki/Wrapped_asymmetric_Laplace_distribution)
  
 
      
      (Wikipedia)
 The [Dirac comb ](https://en.wikipedia.org/wiki/Dirac_comb)of
     period 2 π although not strictly a function, is a limiting form of many
     directional distributions. It is essentially a wrapped [Dirac
     delta function](https://en.wikipedia.org/wiki/Dirac_delta_function). It represents a discrete
     probability distribution concentrated at 2πn — a [degenerate distribution ](https://en.wikipedia.org/wiki/Degenerate_distribution)— but the
     notation treats it as if it were a continuous distribution.
 
      
      (Wikipedia)


 


 ##Supported On
     Semi-infinite Intervals, Usually [0,∞)
 The [Beta prime distribution](https://en.wikipedia.org/wiki/Beta_prime_distribution)
 The [Birnbaum–Saunders distribution](https://en.wikipedia.org/wiki/Birnbaum%E2%80%93Saunders_distribution), also known as the fatigue life
     distribution, is a probability distribution used extensively in
     reliability applications to model failure times.
 The [chi
     distribution ](https://en.wikipedia.org/wiki/Chi_distribution)
 
  The [noncentral chi distribution](https://en.wikipedia.org/wiki/Noncentral_chi_distribution)
 
 The [chi-squared distribution](https://en.wikipedia.org/wiki/Chi-squared_distribution), which is the sum of the squares of n
     independent Gaussian random variables. It is a special case of the Gamma
     distribution, and it is used in [goodness-of-fit ](https://en.wikipedia.org/wiki/Goodness-of-fit)tests
     in [statistics](https://en.wikipedia.org/wiki/Statistics). 
 
  The [inverse-chi-squared distribution](https://en.wikipedia.org/wiki/Inverse-chi-squared_distribution)
  The [noncentral chi-squared distribution](https://en.wikipedia.org/wiki/Noncentral_chi-squared_distribution)
  The [Scaled-inverse-chi-squared
      distribution](https://en.wikipedia.org/wiki/Scaled-inverse-chi-squared_distribution)
 
 The [Dagum
     distribution](https://en.wikipedia.org/wiki/Dagum_distribution)
 The [exponential distribution](https://en.wikipedia.org/wiki/Exponential_distribution), which describes the time between
     consecutive rare random events in a process with no memory.
 The [Exponential-logarithmic distribution](https://en.wikipedia.org/wiki/Exponential-logarithmic_distribution)
 The [F-distribution](https://en.wikipedia.org/wiki/F-distribution), which is the distribution of the
     ratio of two (normalized) chi-squared-distributed random variables, used
     in the [analysis
     of variance](https://en.wikipedia.org/wiki/Analysis_of_variance). It is referred to as the [beta prime distribution ](https://en.wikipedia.org/wiki/Beta_prime_distribution)when it is
     the ratio of two chi-squared variates which are not normalized by dividing
     them by their numbers of degrees of freedom. 
 
  The [noncentral F-distribution](https://en.wikipedia.org/wiki/Noncentral_F-distribution)
 
 [Fisher's z-distribution](https://en.wikipedia.org/wiki/Fisher%27s_z-distribution)
 The [folded normal distribution](https://en.wikipedia.org/wiki/Folded_normal_distribution)
 The [Fréchet
     distribution](https://en.wikipedia.org/wiki/Fr%C3%A9chet_distribution)
 The [Gamma
     distribution](https://en.wikipedia.org/wiki/Gamma_distribution), which describes the time until n
     consecutive rare random events occur in a process with no memory. 
 
  The [Erlang
      distribution](https://en.wikipedia.org/wiki/Erlang_distribution), which is a special case of the
      gamma distribution with integral shape parameter, developed to predict
      waiting times in [queuing systems](https://en.wikipedia.org/wiki/Queuing_systems)
  The [inverse-gamma distribution](https://en.wikipedia.org/wiki/Inverse-gamma_distribution)
 
 The [Generalized gamma distribution](https://en.wikipedia.org/wiki/Generalized_gamma_distribution)
 The [generalized Pareto distribution](https://en.wikipedia.org/wiki/Generalized_Pareto_distribution)
 The [Gamma/Gompertz distribution](https://en.wikipedia.org/wiki/Gamma/Gompertz_distribution)
 The [Gompertz
     distribution](https://en.wikipedia.org/wiki/Gompertz_distribution)
 The [half-normal distribution](https://en.wikipedia.org/wiki/Half-normal_distribution)
 [Hotelling's T-squared distribution](https://en.wikipedia.org/wiki/Hotelling%27s_T-squared_distribution)
 The [inverse Gaussian distribution](https://en.wikipedia.org/wiki/Inverse_Gaussian_distribution), also known as the Wald distribution
 The [Lévy
     distribution](https://en.wikipedia.org/wiki/L%C3%A9vy_distribution)
 The [log-Cauchy distribution](https://en.wikipedia.org/wiki/Log-Cauchy_distribution)
 The [log-Laplace distribution](https://en.wikipedia.org/wiki/Log-Laplace_distribution)
 The [log-logistic distribution](https://en.wikipedia.org/wiki/Log-logistic_distribution)
 The [log-normal distribution](https://en.wikipedia.org/wiki/Log-normal_distribution), describing variables which can be
     modelled as the product of many small independent positive variables.
 The [Lomax
     distribution](https://en.wikipedia.org/wiki/Lomax_distribution)
 The [Mittag-Leffler distribution](https://en.wikipedia.org/wiki/Mittag-Leffler_distribution)
 The [Nakagami
     distribution](https://en.wikipedia.org/wiki/Nakagami_distribution)
 The [Pareto
     distribution](https://en.wikipedia.org/wiki/Pareto_distribution), or "power law"
     distribution, used in the analysis of financial data and critical
     behavior.
 The [Pearson
     Type III distribution](https://en.wikipedia.org/wiki/Pearson_distribution)
 The [Phase-type distribution](https://en.wikipedia.org/wiki/Phase-type_distribution), used in [queueing
     theory](https://en.wikipedia.org/wiki/Queueing_theory)
 The [phased bi-exponential distribution ](https://en.wikipedia.org/w/index.php?title=Phased_bi-exponential_distribution&action=edit&redlink=1)is
     commonly used in [pharmokinetics](https://en.wikipedia.org/wiki/Pharmokinetics)
 The [phased bi-Weibull distribution](https://en.wikipedia.org/w/index.php?title=Phased_bi-Weibull_distribution&action=edit&redlink=1)
 The [Rayleigh
     distribution](https://en.wikipedia.org/wiki/Rayleigh_distribution)
 The [Rayleigh mixture distribution](https://en.wikipedia.org/wiki/Rayleigh_mixture_distribution)
 The [Rice
     distribution](https://en.wikipedia.org/wiki/Rice_distribution)
 The [shifted Gompertz distribution](https://en.wikipedia.org/wiki/Shifted_Gompertz_distribution)
 The [type-2 Gumbel distribution](https://en.wikipedia.org/wiki/Type-2_Gumbel_distribution)
 The [Weibull
     distribution ](https://en.wikipedia.org/wiki/Weibull_distribution)or
     Rosin Rammler distribution, of which the [exponential distribution ](https://en.wikipedia.org/wiki/Exponential_distribution)is a special
     case, is used to model the lifetime of technical devices and is used to
     describe the [particle size distribution ](https://en.wikipedia.org/wiki/Particle_size_distribution)of
     particles generated by grinding, [milling ](https://en.wikipedia.org/wiki/Mill_(grinding))and
     [crushing ](https://en.wikipedia.org/wiki/Crusher)operations.


 


 ##Supported On The
     Whole Real Line
 The [Behrens–Fisher distribution](https://en.wikipedia.org/wiki/Behrens%E2%80%93Fisher_distribution), which arises in the [Behrens–Fisher problem](https://en.wikipedia.org/wiki/Behrens%E2%80%93Fisher_problem).
 The [Cauchy
     distribution](https://en.wikipedia.org/wiki/Cauchy_distribution), an example of a distribution which
     does not have an [expected value ](https://en.wikipedia.org/wiki/Expected_value)or
     a [variance](https://en.wikipedia.org/wiki/Variance). In physics it is usually called a [Lorentzian
     profile](https://en.wikipedia.org/wiki/Lorentzian_function), and is associated with many
     processes, including [resonance ](https://en.wikipedia.org/wiki/Resonance)energy
     distribution, impact and natural [spectral line ](https://en.wikipedia.org/wiki/Spectral_line)broadening
     and quadratic [stark ](https://en.wikipedia.org/wiki/Stark_effect)line
     broadening.
 [Chernoff's distribution](https://en.wikipedia.org/wiki/Chernoff%27s_distribution)
 The [Exponentially modified Gaussian distribution](https://en.wikipedia.org/wiki/Exponentially_modified_Gaussian_distribution), a convolution of a [normal
     distribution ](https://en.wikipedia.org/wiki/Normal_distribution)with
     an [exponential distribution](https://en.wikipedia.org/wiki/Exponential_distribution).
 The [Fisher–Tippett](https://en.wikipedia.org/wiki/Fisher%E2%80%93Tippett_distribution), extreme value, or log-Weibull
     distribution
 [Fisher's z-distribution](https://en.wikipedia.org/wiki/Fisher%27s_z-distribution)
 The [skewed generalized t distribution](https://en.wikipedia.org/wiki/Skewed_generalized_t_distribution)
 The [generalized logistic distribution](https://en.wikipedia.org/wiki/Generalized_logistic_distribution)
 The [generalized normal distribution](https://en.wikipedia.org/wiki/Generalized_normal_distribution)
 The [geometric stable distribution](https://en.wikipedia.org/wiki/Geometric_stable_distribution)
 The [Gumbel
     distribution](https://en.wikipedia.org/wiki/Gumbel_distribution)
 The [Holtsmark distribution](https://en.wikipedia.org/wiki/Holtsmark_distribution), an example of a distribution that
     has a finite expected value but infinite variance.
 The [hyperbolic distribution](https://en.wikipedia.org/wiki/Hyperbolic_distribution)
 The [hyperbolic secant distribution](https://en.wikipedia.org/wiki/Hyperbolic_secant_distribution)
 The [Johnson SU distribution](https://en.wikipedia.org/wiki/Johnson_SU_distribution)
 The [Landau
     distribution](https://en.wikipedia.org/wiki/Landau_distribution)
 The [Laplace
     distribution](https://en.wikipedia.org/wiki/Laplace_distribution)
 The [Lévy skew alpha-stable distribution ](https://en.wikipedia.org/wiki/L%C3%A9vy_skew_alpha-stable_distribution)or
     [stable
     distribution ](https://en.wikipedia.org/wiki/Stable_distribution)is
     a family of distributions often used to characterize financial data and
     critical behavior; the [Cauchy
     distribution](https://en.wikipedia.org/wiki/Cauchy_distribution), [Holtsmark distribution](https://en.wikipedia.org/wiki/Holtsmark_distribution), [Landau
     distribution](https://en.wikipedia.org/wiki/Landau_distribution), [Lévy
     distribution ](https://en.wikipedia.org/wiki/L%C3%A9vy_distribution)and
     [normal
     distribution ](https://en.wikipedia.org/wiki/Normal_distribution)are
     special cases.
 The [Linnik
     distribution](https://en.wikipedia.org/wiki/Linnik_distribution)
 The [logistic
     distribution](https://en.wikipedia.org/wiki/Logistic_distribution)
 The [map-Airy distribution](https://en.wikipedia.org/w/index.php?title=Map-Airy_distribution&action=edit&redlink=1)
 The [normal
     distribution](https://en.wikipedia.org/wiki/Normal_distribution), also called the Gaussian or the
     bell curve. It is ubiquitous in nature and statistics due to the [central
     limit theorem](https://en.wikipedia.org/wiki/Central_limit_theorem): every variable that can be modelled
     as a sum of many small independent, identically distributed variables with
     finite [mean ](https://en.wikipedia.org/wiki/Mean)and [variance ](https://en.wikipedia.org/wiki/Variance)is
     approximately normal.
 The [Normal-exponential-gamma distribution](https://en.wikipedia.org/wiki/Normal-exponential-gamma_distribution)
 The [Normal-inverse Gaussian distribution](https://en.wikipedia.org/wiki/Normal-inverse_Gaussian_distribution)
 The [Pearson Type IV distribution ](https://en.wikipedia.org/w/index.php?title=Pearson_Type_IV_distribution&action=edit&redlink=1)(see
     [Pearson
     distributions](https://en.wikipedia.org/wiki/Pearson_distribution))
 The [skew normal distribution](https://en.wikipedia.org/wiki/Skew_normal_distribution)
 [Student's t-distribution](https://en.wikipedia.org/wiki/Student%27s_t-distribution), useful for estimating unknown means
     of Gaussian populations. 
 
  The [noncentral t-distribution](https://en.wikipedia.org/wiki/Noncentral_t-distribution)
  The [skew t distribution](https://en.wikipedia.org/w/index.php?title=Skew_t_distribution&action=edit&redlink=1)
 
 The [type-1 Gumbel distribution](https://en.wikipedia.org/wiki/Type-1_Gumbel_distribution)
 The [Tracy–Widom distribution](https://en.wikipedia.org/wiki/Tracy%E2%80%93Widom_distribution)
 The [Voigt distribution](https://en.wikipedia.org/wiki/Voigt_profile), or Voigt profile, is the
     convolution of a [normal
     distribution ](https://en.wikipedia.org/wiki/Normal_distribution)and
     a [Cauchy
     distribution](https://en.wikipedia.org/wiki/Cauchy_distribution). It is found in spectroscopy when [spectral line ](https://en.wikipedia.org/wiki/Spectral_line)profiles
     are broadened by a mixture of [Lorentzian ](https://en.wikipedia.org/wiki/Lorentzian_function)and
     [Doppler ](https://en.wikipedia.org/wiki/Doppler_broadening)broadening
     mechanisms.
 The [Gaussian minus exponential distribution ](https://en.wikipedia.org/wiki/Gaussian_minus_exponential_distribution)is
     a convolution of a [normal
     distribution ](https://en.wikipedia.org/wiki/Normal_distribution)with
     (minus) an [exponential distribution](https://en.wikipedia.org/wiki/Exponential_distribution).
 The [Chen distribution](https://en.wikipedia.org/w/index.php?title=Chen_distribution&action=edit&redlink=1).


 


 ##With Variable Support
 The [generalized extreme value distribution ](https://en.wikipedia.org/wiki/Generalized_extreme_value_distribution)has
     a finite upper bound or a finite lower bound depending on what range the
     value of one of the parameters of the distribution is in (or is supported
     on the whole real line for one special value of the parameter
 The [generalized Pareto distribution ](https://en.wikipedia.org/wiki/Generalized_Pareto_distribution)has
     a support which is either bounded below only, or bounded both above and
     below
 The [Tukey lambda distribution ](https://en.wikipedia.org/wiki/Tukey_lambda_distribution)is either
     supported on the whole real line, or on a bounded interval, depending on
     what range the value of one of the parameters of the distribution is in.
 The [Wakeby
     distribution](https://en.wikipedia.org/wiki/Wakeby_distribution)


 


 #Mixed
     Discrete/Continuous Distributions
 The [rectified Gaussian distribution ](https://en.wikipedia.org/wiki/Rectified_Gaussian_distribution)replaces
     negative values from a [normal
     distribution ](https://en.wikipedia.org/wiki/Normal_distribution)with
     a discrete component at zero.
 The [compound
     poisson-gamma or Tweedie distribution ](https://en.wikipedia.org/wiki/Tweedie_distribution)is
     continuous over the strictly positive real numbers, with a mass at zero.


 

#Joint Distributions




 ##Two or More Random Variables On The
     Sample Space
 The [Dirichlet distribution](https://en.wikipedia.org/wiki/Dirichlet_distribution), a generalization of the [beta
     distribution](https://en.wikipedia.org/wiki/Beta_distribution).
 The [Ewens's sampling formula ](https://en.wikipedia.org/wiki/Ewens%27s_sampling_formula)is a
     probability distribution on the set of all [partitions
     of an integer ](https://en.wikipedia.org/wiki/Integer_partition)n,
     arising in [population
     genetics](https://en.wikipedia.org/wiki/Population_genetics).
 The [Balding–Nichols
     model](https://en.wikipedia.org/wiki/Balding%E2%80%93Nichols_model)
 The [multinomial distribution](https://en.wikipedia.org/wiki/Multinomial_distribution), a generalization of the [binomial
     distribution](https://en.wikipedia.org/wiki/Binomial_distribution).
 The [multivariate normal distribution](https://en.wikipedia.org/wiki/Multivariate_normal_distribution), a generalization of the [normal
     distribution](https://en.wikipedia.org/wiki/Normal_distribution).
 The [multivariate t-distribution](https://en.wikipedia.org/wiki/Multivariate_t-distribution), a generalization of the [Student's t-distribution](https://en.wikipedia.org/wiki/Student%27s_t-distribution).
 The [negative multinomial distribution](https://en.wikipedia.org/wiki/Negative_multinomial_distribution), a generalization of the [negative binomial distribution](https://en.wikipedia.org/wiki/Negative_binomial_distribution).
 The [generalized multivariate log-gamma
     distribution](https://en.wikipedia.org/wiki/Generalized_multivariate_log-gamma_distribution)


 


 ##Matrix-Valued Distributions
 The [Wishart
     distribution](https://en.wikipedia.org/wiki/Wishart_distribution)
 The [inverse-Wishart distribution](https://en.wikipedia.org/wiki/Inverse-Wishart_distribution)
 The [matrix normal distribution](https://en.wikipedia.org/wiki/Matrix_normal_distribution)
 The [matrix
     t-distribution](https://en.wikipedia.org/wiki/Matrix_t-distribution)


 


 #Non-numeric
     Distributions
 The [categorical distribution](https://en.wikipedia.org/wiki/Categorical_distribution)



 #Miscellaneous
     Distributions
 The [Cantor
     distribution](https://en.wikipedia.org/wiki/Cantor_distribution)
 The [generalized logistic distribution ](https://en.wikipedia.org/wiki/Generalized_logistic_distribution)family
 The [Pearson
     distribution ](https://en.wikipedia.org/wiki/Pearson_distribution)family
 The [phase-type distribution](https://en.wikipedia.org/wiki/Phase-type_distribution)


 




 


 #Works Cited
 Joyce, D. (2016). Common probability
 distributionsI. Retrieved January 31, 2017, from http://aleph0.clarku.edu/~djoyce/ma218/distributions.pdf
 Minitab. (n.d.). What is a continuous
 distribution? Retrieved January 31, 2017, from
 http://support.minitab.com/en-us/minitab/17/topic-library/basic-statistics-and-graphs/probability-distributions-and-random-data/basics/continuous-distribution/
 Sonnier, R. (n.d.). Kolmogorov Smirnov VBA Code.
 Retrieved January 31, 2017, from
 http://rolfsonnier.typepad.com/blog/2012/07/download-kolmogorov-smirnov-vba-code.html
 Weisstein, E. (2017, January 31). Discrete
 Distribution. Retrieved January 31, 2017, from
 http://mathworld.wolfram.com/DiscreteDistribution.html
 Weisstein, E. (2017, January 31). Wolfram
 Mathworld. Retrieved January 31, 2017, from
 http://mathworld.wolfram.com/BernoulliDistribution.html
 Wikipedia. (n.d.). Dirac comb. Retrieved
 January 31, 2017, from https://en.wikipedia.org/wiki/Dirac_comb
 Wikipedia. (n.d.). Irwin–Hall distribution.
 Retrieved January 31, 2017, from
 https://en.wikipedia.org/wiki/Irwin%E2%80%93Hall_distribution
 Wikipedia. (n.d.). Kent distribution. Retrieved
 January 31, 2017, from https://en.wikipedia.org/wiki/Kent_distribution
 Wikipedia. (n.d.). Kumaraswamy distribution.
 Retrieved January 31, 2017, from
 https://en.wikipedia.org/wiki/Kumaraswamy_distribution
 Wikipedia. (n.d.). List of Probability Distributions.
 Retrieved January 31, 2017, from
 https://en.wikipedia.org/wiki/List_of_probability_distributions
 Wikipedia. (n.d.). Logit-normal distribution.
 Retrieved January 31, 2017, from
 https://en.wikipedia.org/wiki/Logit-normal_distribution
 Wikipedia. (n.d.). Raised cosine distribution.
 Retrieved January 31, 2017, from
 https://en.wikipedia.org/wiki/Raised_cosine_distribution
 Wikipedia. (n.d.). Triangular distribution.
 Retrieved January 31, 2017, from
 https://en.wikipedia.org/wiki/Triangular_distribution
 Wikipedia. (n.d.). Uniform distribution
 (continuous). Retrieved January 31, 2017, from
 https://en.wikipedia.org/wiki/Uniform_distribution_(continuous)
 Wikipedia. (n.d.). von Mises distribution.
 Retrieved January 31, 2017, from https://en.wikipedia.org/wiki/Von_Mises_distribution
 Wikipedia. (n.d.). Wrapped asymmetric Laplace
 distribution. Retrieved January 31, 2017, from
 https://en.wikipedia.org/wiki/Wrapped_asymmetric_Laplace_distribution
 Wikipedia. (n.d.). Wrapped Cauchy distribution.
 Retrieved January 31, 2017, from
 https://en.wikipedia.org/wiki/Wrapped_Cauchy_distribution
 Wikipedia. (n.d.). Wrapped exponential
 distribution. Retrieved January 31, 2017, from
 https://en.wikipedia.org/wiki/Wrapped_exponential_distribution
 Wikipedia. (n.d.). Wrapped normal distribution.
 Retrieved January 31, 2017, from
 https://en.wikipedia.org/wiki/Wrapped_normal_distribution
 Zaiontz, C. (n.d.). Shapiro-Wilk Expanded Test.
 Retrieved January 31, 2017, from
 http://www.real-statistics.com/tests-normality-and-symmetry/statistical-tests-normality-symmetry/shapiro-wilk-expanded-test/
 Zaiontz, C. (n.d.). Shapiro-Wilk Original Test.
 Retrieved January 31, 2017, from
 http://www.real-statistics.com/tests-normality-and-symmetry/statistical-tests-normality-symmetry/shapiro-wilk-test/
  


 

                        

 

